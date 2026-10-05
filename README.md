# Excel Macro Assistant

A Windows desktop app for turning a plain-language task into a VBA macro and saving it in a new `.xlsm` copy of an Excel workbook. It uses desktop Microsoft Excel through COM automation. It does not modify the selected source workbook.

## Requirements

- Windows 10 or 11
- Desktop Microsoft Excel
- Python 3.10 or newer for running from source
- A Gemini API key and a Gemini model name, or OpenAI credentials if selecting OpenAI

## Run from source

1. Install Python, then open PowerShell in this folder.
2. Create an environment and install the dependencies:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   ```

3. Edit `.env` and set `LLM_MODEL` and `GEMINI_API_KEY`. The repository `.env` contains empty placeholders; it does not contain an API key. Keep your real key private. `LLM_PROVIDER` can be left blank to use `config.yaml`, or set to `gemini` or `openai`.
4. Optionally copy `config.example.yaml` to `config.yaml` and set a provider/model. Environment variables take precedence over YAML values. The app also lets you choose Gemini or OpenAI in the provider dropdown.
5. Start the app:

   ```powershell
   python main.py
   ```

The `.env` file is excluded from Git. API keys are read from the selected provider's environment variable (`GEMINI_API_KEY` or `OPENAI_API_KEY`) and are not written to the workbook or logs. The model is supplied by `LLM_MODEL` or `config.yaml`; the source does not select a particular model name. Provider implementations self-register and supply their API-key environment variable, so a new provider module appears in the dropdown without editing the UI or config loader.

## Use

Choose an `.xlsx` or `.xlsm` workbook, describe the macro in plain English, and select **Preview macro**. The app sends the request plus sheet names, used ranges, and up to 50 header cells per sheet to the selected provider. It does not send the workbook's body data. Review the summary and apply the macro when ready.

The new workbook is saved as `Macro Assistant Output/<original-name>.xlsm` beside the source. If that output already exists, the next copy is saved in a timestamped subfolder so earlier generated workbooks are not overwritten. A timestamped copy of the original is stored in `Macro Assistant Output/backups`. The output keeps the original base filename. The app adds only a standard VBA module; it does not add buttons or workbook-open hooks. In Excel, open the new workbook, enable macros if prompted, press `Alt+F8`, select the macro, and choose **Run**.

A static scanner flags common shell, file deletion, registry, and network operations. Flagged code requires an explicit confirmation. This scanner is a warning layer, not a security sandbox or a guarantee that generated VBA is safe. Review generated code and only use workbooks and macros you trust. Failed injection or a detected compile issue can trigger up to two automatic repair requests; the error and workbook metadata are sent to the configured provider for repair.

## Excel VBA access

Excel must allow programmatic access to the VBA project object model. The app checks the current user's `AccessVBOM` setting and includes a confirmation-based helper to set it. Close Excel before applying the change. The setting is stored at:

`HKEY_CURRENT_USER\Software\Microsoft\Office\<version>\Excel\Security\AccessVBOM` (`DWORD` value `1`).

For managed devices, IT can configure **Trust access to the VBA project object model** through Office Group Policy, commonly under **User Configuration > Administrative Templates > Microsoft Excel 2016 > Excel Options > Security > Trust Center**. Policy paths and labels may vary by Office/ADMX version. Organization policy takes precedence; the app will not try to override it. The setting allows programmatic editing of VBA projects and should be enabled only under your organization's security policy.

## Tests and packaging

Run the unit tests without making API calls:

```powershell
python -m pytest
```

The Excel injection integration test skips if Windows, desktop Excel, pywin32, or VBA project access is unavailable.

Build a single-file Windows executable after installing requirements:

```powershell
.\build.bat
```

Place `.env` next to `dist\ExcelMacroAssistant.exe` and fill in the provider key and model. The generated executable is in `dist`. Keep `.env` out of shared or public folders.

## Logs and limitations

Structured logs are written to `%LOCALAPPDATA%\MacroAssistant\logs\app.log`. They contain event names and timestamps, not API keys, prompts, or cell values.

This application requires Windows and desktop Excel and is not intended for server-side use. VBA project access can be disabled by organization policy. Excel's COM interface does not provide a dependable headless compile diagnostic in every installation, so compilation is best-effort. Generated code still needs human review. The context sent to an AI provider includes workbook sheet names, used-range addresses, and header labels, which can themselves contain sensitive information.
