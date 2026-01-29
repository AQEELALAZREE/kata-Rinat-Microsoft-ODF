# GitHub Setup Guide

## Quick Setup (Recommended)

### 1. Initialize Git Repository

```bash
cd "/Users/aqeelalazree/Downloads/kata Rinat 2026"
git init
git add .
git commit -m "Initial commit: Excel Formula Engine documentation and tooling"
```

### 2. Create GitHub Repository

**Option A: Using GitHub CLI (fastest)**
```bash
# Install GitHub CLI if not already installed
# brew install gh

# Authenticate
gh auth login

# Create repository and push
gh repo create excel-formula-engine --public --source=. --remote=origin --push
```

**Option B: Using GitHub Web Interface**

1. Go to https://github.com/new
2. Repository name: `excel-formula-engine`
3. Description: "Microsoft Excel formula engine implementation with comprehensive specs and TDD workflow"
4. Choose Public or Private
5. **Do NOT** initialize with README, .gitignore, or license (we already have these)
6. Click "Create repository"

### 3. Connect and Push (if using Option B)

```bash
# Add remote
git remote add origin https://github.com/YOUR_USERNAME/excel-formula-engine.git

# Push to GitHub
git branch -M main
git push -u origin main
```

## Detailed Setup

### Step 1: Initialize Local Repository

```bash
cd "/Users/aqeelalazree/Downloads/kata Rinat 2026"

# Initialize git
git init

# Check status
git status
```

### Step 2: Stage All Files

```bash
# Add all files (respects .gitignore)
git add .

# Verify what will be committed
git status
```

### Step 3: Create Initial Commit

```bash
git commit -m "Initial commit: Excel Formula Engine

- Complete architecture documentation
- Technology stack recommendations
- Parser, dependency graph, and function library specs
- Testing strategy and implementation guide
- ODF formula specifications with Excel validation tools
- TDD workflow for incremental implementation
- Code review checklists and examples"
```

### Step 4: Create GitHub Repository

Choose one of these methods:

#### Method 1: GitHub CLI (Easiest)

```bash
# Create and push in one command
gh repo create excel-formula-engine \
  --public \
  --description "Excel formula engine with comprehensive specs and TDD workflow" \
  --source=. \
  --remote=origin \
  --push
```

#### Method 2: GitHub Web + Manual Push

1. **Create on GitHub**:
   - Go to https://github.com/new
   - Name: `excel-formula-engine`
   - Description: "Excel formula engine with comprehensive specs and TDD workflow"
   - Public/Private: Your choice
   - **Important**: Leave all checkboxes unchecked (no README, .gitignore, license)
   - Click "Create repository"

2. **Connect and Push**:
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/excel-formula-engine.git
   git branch -M main
   git push -u origin main
   ```

### Step 5: Verify Upload

```bash
# Check remote
git remote -v

# View on GitHub
gh repo view --web
# Or manually visit: https://github.com/YOUR_USERNAME/excel-formula-engine
```

## Repository Settings (Optional)

### Add Topics

On GitHub repository page:
1. Click "⚙️ Settings" or the gear icon next to "About"
2. Add topics: `excel`, `formula-engine`, `parser`, `spreadsheet`, `tdd`, `odf`

### Set Up Branch Protection

1. Go to Settings → Branches
2. Add rule for `main` branch:
   - ✅ Require pull request reviews
   - ✅ Require status checks to pass

### Enable GitHub Actions (for CI/CD)

Create `.github/workflows/test.yml`:

```yaml
name: Tests

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      - name: Install dependencies
        run: |
          pip install -r requirements.txt
      - name: Run tests
        run: pytest tests/ -v --cov
```

## Recommended Repository Structure

Your repository is already well-structured:

```
excel-formula-engine/
├── .gitignore              ✅ Created
├── README.md               ✅ Exists
├── requirements.txt        ✅ Exists
├── docs/                   ✅ Documentation
├── specs/                  ✅ Specifications
├── scripts/                ✅ Validation tools
├── test-cases/             ✅ Test scenarios
├── examples/               ✅ Usage examples
├── review-checklists/      ✅ QA guidelines
└── .agent/workflows/       ✅ Agent workflows
```

## Common Issues

### Issue: "fatal: not a git repository"
**Solution**: Make sure you're in the correct directory
```bash
cd "/Users/aqeelalazree/Downloads/kata Rinat 2026"
git init
```

### Issue: "remote origin already exists"
**Solution**: Update the remote URL
```bash
git remote set-url origin https://github.com/YOUR_USERNAME/excel-formula-engine.git
```

### Issue: Large files rejected
**Solution**: Check .gitignore is working
```bash
git status  # Should not show large Excel files
```

## Next Steps After Upload

1. **Add README badges**:
   ```markdown
   ![Python](https://img.shields.io/badge/python-3.11+-blue.svg)
   ![License](https://img.shields.io/badge/license-MIT-green.svg)
   ```

2. **Create issues** for implementation tasks

3. **Set up project board** for tracking progress

4. **Invite collaborators** if working with a team

5. **Enable Discussions** for Q&A

## Useful Git Commands

```bash
# Check status
git status

# View commit history
git log --oneline

# Create new branch
git checkout -b feature/parser-implementation

# Push branch
git push -u origin feature/parser-implementation

# Pull latest changes
git pull origin main
```

## Repository URL

After setup, your repository will be at:
```
https://github.com/YOUR_USERNAME/excel-formula-engine
```

Replace `YOUR_USERNAME` with your GitHub username.
