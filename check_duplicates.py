import ast
import sys

def check_duplicates(filename):
    with open(filename, "r", encoding="utf-8") as f:
        tree = ast.parse(f.read())
    
    routes = []
    funcs = []
    for node in tree.body:
        if isinstance(node, ast.AsyncFunctionDef) or isinstance(node, ast.FunctionDef):
            funcs.append(node.name)
            for decorator in node.decorator_list:
                if isinstance(decorator, ast.Call) and hasattr(decorator.func, 'attr'):
                    if decorator.func.attr in ['get', 'post', 'put', 'delete']:
                        if isinstance(decorator.args[0], ast.Constant):
                            routes.append((decorator.func.attr, decorator.args[0].value))
    
    from collections import Counter
    dup_funcs = [f for f, c in Counter(funcs).items() if c > 1]
    dup_routes = [r for r, c in Counter(routes).items() if c > 1]
    
    if dup_funcs:
        print(f"Duplicate Functions: {dup_funcs}")
    if dup_routes:
        print(f"Duplicate Routes: {dup_routes}")
    if not dup_funcs and not dup_routes:
        print("No duplicate functions or routes found.")

if __name__ == "__main__":
    check_duplicates("main.py")
