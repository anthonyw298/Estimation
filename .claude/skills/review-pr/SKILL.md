---
name: review-pr
description: Review a pull request for code quality, patterns, and issues
disable-model-invocation: true
context: fork
agent: Explore
allowed-tools: Bash(gh *), Read, Grep, Glob
---

# Review Pull Request

Review PR $ARGUMENTS for code quality and potential issues.

## Context

- PR diff: !`gh pr diff $ARGUMENTS`
- PR details: !`gh pr view $ARGUMENTS`
- Changed files: !`gh pr diff $ARGUMENTS --name-only`

## Review Checklist

1. **Correctness**: Does the code do what it claims?
2. **TypeScript**: Are types correct and specific (no unnecessary `any`)?
3. **React patterns**: Server vs Client components used correctly? React 19 APIs?
4. **Performance**: Any unnecessary re-renders, missing memoization, or N+1 queries?
5. **Security**: No exposed secrets, SQL injection, or XSS vulnerabilities?
6. **Style**: Follows project conventions (Tailwind, `cn()`, `@/` imports)?
7. **Database**: Supabase queries handle errors and use proper filters?

## Output Format

For each file with findings, list:
- **File**: path
- **Line**: number
- **Severity**: critical / warning / suggestion
- **Issue**: description
- **Fix**: recommended change

End with an overall summary and approval recommendation.
