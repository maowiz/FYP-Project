# Git Push Guide - What Will Be Included/Excluded

## Overview

The `.gitignore` file has been updated to exclude large files and unnecessary dependencies while keeping all essential source code.

---

## ✅ WHAT WILL BE PUSHED (Included)

### Backend (Python)

- ✅ All `.py` files (Python source code)
- ✅ `main.py`, `server.py`, `command_handler.py`, etc.
- ✅ `os_command_handler.py`, `file_command_handler.py`, `general_command_handler.py`
- ✅ `hybrid_processor.py`, `intent_classifier.py`
- ✅ `speech.py`, `voice_recognition.py`
- ✅ `browser_commands.py`, `os_management.py`, `file_management.py`
- ✅ All other Python source files

### Frontend (React/TypeScript)

- ✅ `src/` folder (all source code)
- ✅ `public/` folder (static assets)
- ✅ `index.html` (HTML entry point)
- ✅ `package.json` (dependencies list)
- ✅ `tsconfig.json` (TypeScript config)
- ✅ `vite.config.ts` (Vite config)
- ✅ `README.md` (documentation)

### Configuration & Documentation

- ✅ `.gitignore` (this file)
- ✅ `config.json` (if not sensitive)
- ✅ `ALL_COMMANDS.md`
- ✅ `README.md` files

### Auth Models

- ✅ `auth/models/face_model.yml` (YAML config file)

---

## ❌ WHAT WILL NOT BE PUSHED (Excluded)

### Large Model Files

- ❌ `*.gguf` files (Qwen model, etc.)
- ❌ `qwen2.5-0.5b-instruct-q4_k_m.gguf`
- ❌ `*.bin` files (model binaries)
- ❌ `*.safetensors` files

### Node.js Dependencies

- ❌ `frontend advance/node_modules/` (install with `npm install`)
- ❌ `package-lock.json` (will be regenerated)
- ❌ `yarn.lock`

### Build Artifacts

- ❌ `frontend advance/dist/` (rebuild with `npm run build`)
- ❌ `frontend advance/build/`
- ❌ `frontend advance/.vite/`
- ❌ `frontend advance/.next/`
- ❌ `dist/` (Python)
- ❌ `build/` (Python)

### Cache & Temporary Files

- ❌ `__pycache__/` (Python cache)
- ❌ `*.pyc`, `*.pyo` (Python bytecode)
- ❌ `frontend advance/.cache/`
- ❌ `frontend advance/.turbo/`
- ❌ `*.log` (log files)
- ❌ `*.tmp` (temporary files)

### ML/AI Models & Data

- ❌ `*.pkl`, `*.pickle` (pickle files)
- ❌ `face_encodings.pkl`
- ❌ `labels.pkl`
- ❌ `face_database.pkl`
- ❌ `intent_classifier.pkl`
- ❌ `auth/dataset/` (large face dataset)
- ❌ `dataset/` folder

### Sensitive Files

- ❌ `credentials.json` (Google API credentials)
- ❌ `token.pickle` (Gmail token)
- ❌ `.env` files (environment variables)
- ❌ `*.pem`, `*.key`, `*.crt` (certificates)

### IDE & OS Files

- ❌ `.vscode/` (VS Code settings)
- ❌ `.idea/` (IntelliJ settings)
- ❌ `.DS_Store` (macOS)
- ❌ `Thumbs.db` (Windows)

---

## 📋 Steps to Push Code

### 1. Check what will be pushed

```bash
git status
```

### 2. Add all files (respecting .gitignore)

```bash
git add .
```

### 3. Commit with a message

```bash
git commit -m "Add voice assistant with command feedback and caching fixes"
```

### 4. Push to repository

```bash
git push origin main
```

---

## 🔄 After Cloning (For Others)

When someone clones your repository, they'll need to:

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Frontend dependencies

```bash
cd "frontend advance"
npm install
cd ..
```

### 3. Download the Qwen model (if needed)

```bash
# Download from Hugging Face or your source
# Place in the models/ directory
```

### 4. Set up credentials

```bash
# Create .env file with your API keys
# Create credentials.json for Gmail API
```

---

## ⚠️ Important Notes

1. **Qwen Model**: The `.gguf` file is excluded because it's too large (>500MB). Users can download it separately.

2. **node_modules**: Excluded because it's huge. Users run `npm install` to regenerate it.

3. **Credentials**: Never pushed to git. Users must set up their own `.env` and `credentials.json`.

4. **Frontend Build**: The `dist/` folder is excluded. Users rebuild with `npm run build`.

5. **Python Cache**: `__pycache__/` is excluded. It's regenerated automatically.

---

## 📊 Repository Size

With these exclusions:

- **Without exclusions**: ~2-3 GB (models, node_modules, cache)
- **With exclusions**: ~50-100 MB (source code only)

This makes cloning and collaboration much faster! ✨
