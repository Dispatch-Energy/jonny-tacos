# Git Commands Quick Reference

## Daily Workflow

```bash
# Start new feature
git checkout develop
git pull origin develop
git checkout -b feature/your-feature-name

# Make changes and commit
git add .
git commit -m "feat: your feature description"

# Push changes
git push origin feature/your-feature-name

# Create pull request on GitHub/Azure DevOps
```

## Common Commands

```bash
# Check status
git status

# View commit history
git log --oneline --graph --all

# Stash changes
git stash
git stash pop

# Undo last commit (keep changes)
git reset --soft HEAD~1

# Sync with remote
git fetch --all
git pull origin main

# Clean up branches
git branch -d feature/branch-name
git remote prune origin
```

## Commit Message Format

```
<type>(<scope>): <subject>

<body>

<footer>
```

Types:
- feat: New feature
- fix: Bug fix
- docs: Documentation
- style: Code style
- refactor: Code refactoring
- test: Testing
- chore: Maintenance
- deploy: Deployment

Examples:
```
feat(quickbase): Add ticket priority auto-assignment
fix(teams): Resolve adaptive card rendering issue
docs(readme): Update deployment instructions
```
