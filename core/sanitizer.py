"""
AST-based static code analyzer to sanitize and validate AI-generated Python code
before execution. Prevents RCE, unauthorized module imports, and system tampering.
"""

import ast
from typing import Set

ALLOWED_ROOT_MODULES: Set[str] = {
    "openpyxl",
    "xlsxwriter",
    "pandas",
    "numpy",
    "datetime",
    "math",
    "random",
    "decimal",
    "itertools",
    "collections",
    "re",
    "json",
    "string",
    "core"
}

FORBIDDEN_NAMES: Set[str] = {
    "eval",
    "exec",
    "compile",
    "open",
    "input",
    "globals",
    "locals",
    "vars",
    "dir",
    "getattr",
    "setattr",
    "delattr",
    "__import__",
    "breakpoint",
    "memoryview",
    "exit",
    "quit"
}

FORBIDDEN_ATTRIBUTES: Set[str] = {
    "__subclasses__",
    "__bases__",
    "__base__",
    "__mro__",
    "__globals__",
    "__code__",
    "__builtins__",
    "__class__",
    "__qualname__",
    "__dict__",
    "__func__",
    "func_globals",
    "func_code",
    "gi_frame",
    "f_globals",
    "f_locals",
    "f_code",
    "cr_frame"
}


class SecurityError(Exception):
    """Raised when generated code violates security policies."""
    pass


class CodeSanitizer(ast.NodeVisitor):
    def __init__(self):
        self.found_generate_excel_function = False

    def visit_FunctionDef(self, node: ast.FunctionDef):
        if node.name == "generate_excel":
            self.found_generate_excel_function = True
        self.generic_visit(node)

    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef):
        if node.name == "generate_excel":
            self.found_generate_excel_function = True
        self.generic_visit(node)

    def visit_Import(self, node: ast.Import):
        for alias in node.names:
            root_module = alias.name.split('.')[0]
            if root_module not in ALLOWED_ROOT_MODULES:
                raise SecurityError(
                    f"Security Violation: Unauthorized module import '{alias.name}'. "
                    f"Allowed modules are: {', '.join(sorted(ALLOWED_ROOT_MODULES))}"
                )
        self.generic_visit(node)

    def visit_ImportFrom(self, node: ast.ImportFrom):
        if node.module:
            root_module = node.module.split('.')[0]
            if root_module not in ALLOWED_ROOT_MODULES:
                raise SecurityError(
                    f"Security Violation: Unauthorized import from '{node.module}'. "
                    f"Allowed modules are: {', '.join(sorted(ALLOWED_ROOT_MODULES))}"
                )
        self.generic_visit(node)

    def visit_Name(self, node: ast.Name):
        if node.id in FORBIDDEN_NAMES:
            raise SecurityError(f"Security Violation: Forbidden function/variable '{node.id}'")
        self.generic_visit(node)

    def visit_Attribute(self, node: ast.Attribute):
        if node.attr in FORBIDDEN_ATTRIBUTES or (node.attr.startswith("__") and node.attr.endswith("__")):
            raise SecurityError(f"Security Violation: Forbidden attribute access '{node.attr}'")
        self.generic_visit(node)

    def visit_Call(self, node: ast.Call):
        if isinstance(node.func, ast.Name):
            if node.func.id in FORBIDDEN_NAMES:
                raise SecurityError(f"Security Violation: Forbidden call '{node.func.id}()'")
        self.generic_visit(node)


def sanitize_python_code(code_str: str, require_function: bool = True) -> ast.AST:
    """
    Parses and verifies Python code against security AST rules.
    Returns the parsed AST tree if valid, or raises SecurityError.
    """
    if not code_str or not code_str.strip():
        raise SecurityError("Generated code is empty.")

    try:
        tree = ast.parse(code_str)
    except SyntaxError as e:
        raise SecurityError(f"Python Syntax Error: {e.msg} on line {e.lineno}")

    sanitizer = CodeSanitizer()
    sanitizer.visit(tree)

    if require_function and not sanitizer.found_generate_excel_function:
        raise SecurityError(
            "Contract Violation: The code must define 'def generate_excel(output_path: str):' "
            "as its entry point."
        )

    return tree
