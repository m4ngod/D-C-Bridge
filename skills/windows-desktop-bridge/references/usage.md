# Usage

## Run

```powershell
D:\Users\AIGOD\AppData\Local\Programs\Python311\python.exe D:\Users\AIGOD\.openclaw\workspace\skills\windows-desktop-bridge\scripts\desktop_bridge.py
```

## Test

```powershell
curl http://127.0.0.1:8765/health
curl http://127.0.0.1:8765/windows
curl http://127.0.0.1:8765/foreground
```

## Example actions

Wait for PyCharm to become foreground (3s interval, 5 checks):

```powershell
curl -X POST http://127.0.0.1:8765/wait-foreground -H "Content-Type: application/json" -d '{"title":"UTI-STOCKSIM","intervalSeconds":3,"maxChecks":5}'
```

Launch PyCharm:

```powershell
curl -X POST http://127.0.0.1:8765/launch -H "Content-Type: application/json" -d '{"path":"D:\\Users\\AIGOD\\AppData\\Local\\Programs\\PyCharm Community 2025.2.6\\bin\\pycharm64.exe"}'
```

Activate a window by title substring:

```powershell
curl -X POST http://127.0.0.1:8765/activate -H "Content-Type: application/json" -d '{"title":"PyCharm"}'
```
