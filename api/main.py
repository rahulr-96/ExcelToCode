from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess, tempfile, os
import uuid

app = FastAPI()

@app.post("/generate-csharp")
async def generate_csharp(
    file: UploadFile = File(...),
    cell_address: str = Form(None)
):
    id = uuid.uuid4()
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, id)
        yaml_path = os.path.join(tmpdir, id+"_yaml"+".yaml")
        csharp_path = os.path.join(tmpdir, id+"_output.cs")
        json_map_path = os.path.join(tmpdir, id+"_varmap.json")
        final_csharp_path = os.path.join(id+"_final.cs")

        # Save uploaded Excel file
        with open(excel_path, "wb") as f:
            f.write(await file.read())

        # Step 1: Excel → YAML
        cmd = ["python", "../yaml-generator/generator.py", excel_path, yaml_path]
        if cell_address:
            cmd.append(cell_address)
        subprocess.run(cmd, check=True)

        # Step 2a: YAML → C#
        subprocess.run(
            ["node", "../code-refactor-tool/index.js", yaml_path, csharp_path],
            check=True
        )

        # Step 2b: Generate variable map JSON
        subprocess.run(
            ["node", "../code-refactor-tool/variable-name-generator.js", csharp_path, json_map_path],
            check=True
        )

        # Step 2c: Apply refactor → Final C#
        subprocess.run(
            ["node", "../code-refactor-tool/class-refactor.js", csharp_path, json_map_path, final_csharp_path],
            check=True
        )

        # Return final C# file as download
        return FileResponse(
            final_csharp_path,
            media_type="text/plain",
            filename="GeneratedCode.cs"
        )
