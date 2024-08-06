**打包**

```bash
pyinstaller --onefile --noconsole --hidden-import=babel.numbers --add-data "utils;utils" app.py
```

