# Contributing to CM360 Audit System

Thank you for your interest in contributing to the CM360 Audit System! This document provides guidelines for contributing to this project.

## Code of Conduct

This project adheres to a code of conduct adapted from the [Contributor Covenant](https://www.contributor-covenant.org/). By participating, you are expected to uphold this code.

## How to Contribute

### Reporting Bugs

Before creating bug reports, please check the existing issues to avoid duplicates. When creating a bug report, please use the bug report template and include:

- A clear description of the issue
- Steps to reproduce the problem
- Expected vs actual behavior
- Your environment details (config names, browser, etc.)
- Apps Script execution logs if available

### Suggesting Features

Feature requests are welcome! Please use the feature request template and include:

- A clear description of the problem you're trying to solve
- Your proposed solution
- Use cases and benefits
- Any alternative approaches you've considered

### Development Setup

1. **Prerequisites**
   - Node.js 14+ for clasp
   - Google account with Apps Script access
   - Access to Google Drive and Gmail APIs

2. **Local Development**
   ```bash
   # Clone the repository
   git clone https://github.com/evan-schneider/cm360-audit-system.git
   cd cm360-audit-system
   
   # Install dependencies
   npm install
   
   # Login to clasp (one-time setup)
   npm run login
   
   # Pull latest from Apps Script
   npm run pull
   ```

3. **Making Changes**
   - Create a feature branch: `git checkout -b feature/your-feature-name`
   - Make your changes in the appropriate files
   - Test thoroughly in the Apps Script environment
   - Follow the existing code style and patterns

### Pull Request Process

1. **Before Submitting**
   - Test your changes with multiple configs
   - Verify Gmail/Drive integration works
   - Check that existing functionality isn't broken
   - Update documentation if needed

2. **PR Requirements**
   - Clear description of changes
   - Reference any related issues
   - Include testing notes
   - Follow semantic commit conventions

3. **Review Process**
   - Maintain discussion in PR comments
   - Address reviewer feedback promptly
   - Keep commits focused and atomic

## Development Guidelines

### Code Style

- Use meaningful variable names
- Add comments for complex business logic
- Follow Apps Script best practices
- Keep functions focused and modular

### Testing

- Test with multiple audit configurations
- Verify email delivery and formatting
- Check Drive folder creation and permissions
- Test error handling paths

### Documentation

- Update README.md for significant changes
- Add inline comments for complex algorithms
- Document new configuration options
- Include examples where helpful

## Project Structure

```
├── Code.js                 # Main Apps Script logic
├── ConfigPicker.html       # Configuration selection UI
├── Dashboard.html          # Status dashboard
├── ButtonsSidebar.html     # Admin controls sidebar
├── AdminRefreshPrompt.html # Admin refresh dialog
├── appsscript.json        # Apps Script manifest
├── package.json           # Node.js dependencies
├── README.md              # Project documentation
└── .github/               # GitHub templates and workflows
```

## Deployment

Changes are deployed through clasp:

```bash
# Deploy to Apps Script
npm run deploy

# Check deployment logs
npm run logs
```

## Questions?

If you have questions about contributing, please:

1. Check existing issues and discussions
2. Create a new issue with your question
3. Tag it with the "question" label

## Recognition

Contributors will be recognized in the project documentation. Thank you for helping make this project better!