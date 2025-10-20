import re
import os

def ts_to_py(ts_content: str) -> str:
    # 去除类型声明如 `: DictionaryResource[]`
    ts_content = re.sub(r':\s*DictionaryResource$$\d*\$$', '', ts_content)

    # 替换 const 变量名为 Python 风格
    ts_content = re.sub(r'const\s+([a-zA-Z0-9_]+)', r'\1 =', ts_content)

    # 将 JavaScript 对象语法转为 Python 字典语法
    lines = []
    for line in ts_content.splitlines():
        line = line.strip()

        # 忽略分号和空行
        if line.endswith(";"):
            line = line[:-1]

        # 处理 key: value → "key": value
        line = re.sub(r'^(\s*)([a-zA-Z0-9]+):', r'\1"\2":', line)

        # 特别处理驼峰命名转蛇形命名，如 languageCategory → language_category
        line = re.sub(r'"(language)([A-Z][a-z]+)(.*?)"', r'"\1_\2\3"', line)

        # 替换数组为 Python 列表
        line = line.replace("[", "[").replace("]", "]")

        lines.append(line)

    return "\n".join(lines)

def convert_ts_file_to_py(input_path: str, output_dir: str):
    with open(input_path, "r", encoding="utf-8") as f:
        ts_data = f.read()

    py_data = ts_to_py(ts_data)

    filename = os.path.splitext(os.path.basename(input_path))[0] + ".py"
    output_path = os.path.join(output_dir, filename)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(py_data)

    print(f"✅ 已保存到 {output_path}")

if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("用法: python ts_to_py_converter.py <输入.ts文件路径> <输出目录>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_folder = sys.argv[2]

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    convert_ts_file_to_py(input_file, output_folder)