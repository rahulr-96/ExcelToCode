const fs = require("fs");
const path = require("path");

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceInComments(csCode, key, newName) {
  const keyRegex = new RegExp(escapeRegExp(key), "g");
  const newNameRegex = new RegExp(escapeRegExp(newName), "g");

  return csCode
    .split("\n")
    .map(line => {
      if (line.trim().startsWith("//")) {
        // Inside a comment line → reverse replace newName → key
        return line.replace(newNameRegex, key);
      } else {
        // Normal code line → do the original replace key → newName
        return line.replace(keyRegex, newName);
      }
    })
    .join("\n");
}

/**
 * Replaces all field names in a C# file using a JSON mapping,
 * and adds a comment above each replaced field with the original sheet address.
 */
function replaceFieldsWithComments(csFilePath, jsonFilePath, outputFilePath) {
  let csCode = fs.readFileSync(csFilePath, "utf-8");
  const mapping = JSON.parse(fs.readFileSync(jsonFilePath, "utf-8"));

  // Sort keys by length (descending) to avoid partial replacements
  const keys = Object.keys(mapping).sort((a, b) => b.length - a.length);

  for (const key of keys) {
    const newName = sanitizeIdentifier(mapping[key]);

    // Regex to capture the full field declaration line
    const regex = new RegExp(`(\\s*(?:double|int|float|decimal|string)\\s+)${key}(\\s*=.*;)`, "g");

    csCode = csCode.replace(regex, (match, typePart, restPart) => {
          // Capture indentation from the start of typePart
      const indent = typePart.match(/^\s*/)?.[0] ?? "";
      return `\n${indent}// ${key}${typePart}${newName}${restPart}`;
    });

    // Also replace references to the old name elsewhere in the file ignore comments
    csCode = replaceInComments(csCode, key, newName);
  }

  fs.writeFileSync(outputFilePath, csCode, "utf-8");
  console.log(`Search & replace with comments complete. Output written to ${outputFilePath}`);
}

/**
 * Sanitize to make valid C# identifiers
 */
function sanitizeIdentifier(str) {
  return str
    .replace(/[^a-zA-Z0-9_]/g, "_") // only alphanumeric + underscore
    .replace(/^(\d)/, "_$1"); // prefix underscore if starts with number
}

function main() {
  const args = process.argv.slice(2);

  if (args.length < 3) {
    console.error("Usage: node class-refactor.js <input_cs> <input_json> <output_cs>");
    process.exit(1);
  }

  const csFilePath = path.resolve(args[0]);
  const jsonFilePath = path.resolve(args[1]);
  const outputFilePath = path.resolve(args[2]);

  replaceFieldsWithComments(csFilePath, jsonFilePath, outputFilePath);
}

main();