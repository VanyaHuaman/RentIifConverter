# Tests

Run the regression tests from the repo root:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-RegressionTests.ps1
```

The tests use fixture workbooks and compare generated IIF files against expected text output.
