const fs = require("fs");
const yaml = require("js-yaml");

// ---------------- Helper Functions for range expansion ----------------
function colToNum(col) {
    let num = 0;
    for (const c of col) {
        num = num * 26 + (c.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1);
    }
    return num;
}

function numToCol(n) {
    let col = "";
    while (n > 0) {
        const rem = (n - 1) % 26;
        col = String.fromCharCode(rem + "A".charCodeAt(0)) + col;
        n = Math.floor((n - 1) / 26);
    }
    return col;
}

function expandRangeList(sheet, colStart, rowStart, colEnd, rowEnd) {
    const startColNum = colToNum(colStart);
    const endColNum = colToNum(colEnd);
    const parts = [];
    for (let r = parseInt(rowStart); r <= parseInt(rowEnd); r++) {
        for (let c = startColNum; c <= endColNum; c++) {
            parts.push(`${sheet}_${numToCol(c)}${r}`);
        }
    }
    return parts.join(", ");
}

// ---------------- Main converter ----------------
function yamlToCSharp(yamlFile, outputFile = "Formulas.cs") {
    const data = yaml.load(fs.readFileSync(yamlFile, "utf8"));
    const cellMap = data.cell_map || {};

    const lines = [
        "using System;",
        "using System.Linq;",
        "",
        "class ExcelFormulas {",
        "    public void Compute() {"
    ];

    for (const [cell, formula] of Object.entries(cellMap)) {
        const [sheet, addr] = cell.split("!");
        const varName = `${sheet}_${addr}`;

        if (typeof formula === "number") {
            lines.push(`        double ${varName} = ${formula};`);
        } else if (typeof formula === "string" && formula.startsWith("=")) {
            let expr = formula.slice(1); // remove '='

            // Replace _C_("sheet!cell") with sheet_cell
            expr = expr.replace(/_C_\("([^"]+)!([^"]+)"\)/g, (_, s, a) => `${s}_${a}`);

            // Map functions
            const funcMap = {
                sum_: "sum_",
                "npv(": "npv_",
                average_: "average_",
                max_: "max_",
                min_: "min_",
                "if_": "if_",
                "and_": "and_",
                "or_": "or_"
            };

            for (const [prefix, name] of Object.entries(funcMap)) {
                if (expr.startsWith(prefix)) {
                    // get arguments inside parentheses
                    let argsStr = expr.slice(prefix.length).replace(/^\(|\)$/g, "");

                    // replace _R_("Sheet!A1:B2") with comma-separated list
                    argsStr = argsStr.replace(/_R_\("([^"]+)!([A-Z]+)(\d+):([A-Z]+)(\d+)"\)/g,
                        (_, s, c1, r1, c2, r2) => expandRangeList(s, c1, r1, c2, r2)
                    );

                    expr = `${name}(${argsStr})`;
                    break;
                }
            }

            lines.push(`        double ${varName} = ${expr};`);
        }
    }

    lines.push("    }");
    lines.push("}");
    
    fs.writeFileSync(outputFile, lines.join("\n"), "utf8");
}

// ---------------- Usage ----------------
function main() {
  const args = process.argv.slice(2);

  if (args.length < 2) {
    console.error("Usage: node index.js <input_yaml> <output_cs>");
    process.exit(1);
  }

  const inputYaml = args[0];
  const outputCSharp = args[1];

  yamlToCSharp(inputYaml, outputCSharp);

  console.log(`C# file generated at ${outputCSharp}`);
}

main();
