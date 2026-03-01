# ASP4 - Python-based runtime for Classic ASP/VBScript
ASP4 is a Python-based runtime that executes Classic ASP (VBScript) pages. It provides compatibility with most VBScript built-in functions, the Classic ASP object model (Request, Response, Session, Application, Server), and various COM components commonly used in legacy ASP applications.

---

# Running ASP4

Start the built-in server:

```bash
python -m ASP4/server.py [host] [port] [docroot]
```

Example:
```bash
python ASP4/server.py 0.0.0.0 8080 www
```

Browse to:
```bash
http://localhost:8080/ or http://127.0.0.1:8080/
```

The server will serve both static files and .asp pages from the specified document root.
