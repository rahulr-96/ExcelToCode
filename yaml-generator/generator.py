import sys
from pycel import ExcelCompiler

def main():
    if len(sys.argv) < 3:
        print("Usage: python generator.py <input_excel> <output_yaml> [cell_address]")
        sys.exit(1)

    input_excel = sys.argv[1]
    output_yaml = sys.argv[2]
    cell_address = sys.argv[3] if len(sys.argv) > 3 else None

    excel_compiler = ExcelCompiler(input_excel)

    if cell_address:
        try:
            result = excel_compiler.evaluate(cell_address)
            print(f"Evaluation of {cell_address}: {result}")
        except Exception as e:
            print(f"Failed to evaluate {cell_address}: {e}")

    # Export workbook formulas to YAML
    excel_compiler.to_file(file_types="yaml")
    print(f"YAML generated at {output_yaml}")


if __name__ == "__main__":
    main()
