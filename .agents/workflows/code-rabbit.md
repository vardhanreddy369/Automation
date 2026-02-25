---
description: CodeRabbit - AI Code Review Simulator
---
# CodeRabbit - AI Code Review Simulator

This workflow turns me into an automated Pull Request code reviewer looking for bugs, edge cases, and architectural smells.

**Steps for the Agent:**
1. **Analyze Diff:** If the user points to a file or asks to check uncommitted changes, I will automatically execute a `git diff` or read the entire target file to see what has changed.
2. **High-Level Summary:** I will provide a clear, high-level summary paragraph of the structural changes detected.
3. **Inline Feedback (Code Smells & Bugs):** I will perform a rigorous review. I will output specific inline feedback calling out potential code smells, bugs, variable misnames, logic issues, and performance bottlenecks line-by-line.
4. **Actionable Patches:** For every piece of feedback, I will proactively provide a ready-to-use structural patch format (code blocks) acting just like CodeRabbit's "one-click" PR suggestions so the user can trivially apply the fixes.
