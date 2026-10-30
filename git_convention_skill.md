---
name: git-convention
description: Apply safe, consistent Conventional Commit messages and push practices when the user asks an AI agent to commit or push repository changes.
---

# Git Commit and Push Convention

Use these instructions whenever the user asks to commit, push, or prepare changes for Git.

## Respect the repository and user intent

- Follow repository-specific instructions first (`AGENTS.md`, `CONTRIBUTING.md`, or an established commit convention).
- A request to commit does not authorize a push. Push only when the user explicitly asks for it or clearly requests the complete commit-and-push workflow.
- Preserve unrelated user changes. Stage only files that belong to the requested work.
- Never commit secrets, `.env` files, credentials, local virtual environments, caches, generated build output, or large binaries unless the user explicitly intends to version them.
- Do not amend, rebase, rewrite history, force-push, or delete branches unless the user explicitly requests that operation.
- If a force-push is explicitly required, prefer `--force-with-lease` and verify the target branch first.

## Inspect before committing

Before creating a commit:

1. Check the current branch, upstream, and working-tree status.
2. Review the relevant diff, including staged changes.
3. Check for accidental files, credentials, debug code, and unrelated edits.
4. Run relevant tests or validation in proportion to the change.
5. Run a whitespace/error check such as `git diff --check` when available.

If the repository contains unrelated changes, do not discard or include them. Stage the intended paths explicitly.

## Commit message format

Use Conventional Commits with a lowercase type:

```text
type(optional-scope): concise imperative summary
```

Use an imperative summary that describes the outcome. Keep it specific and avoid a trailing period.

### Types

- `feat`: add a new user-visible capability
- `fix`: correct a bug or incorrect behavior
- `refactor`: improve code structure without changing intended behavior
- `perf`: improve performance
- `docs`: change documentation only
- `test`: add or update tests only
- `build`: change build tooling, packaging, or dependencies
- `ci`: change CI/CD configuration
- `style`: change formatting without changing behavior
- `chore`: perform maintenance that fits no more specific type
- `revert`: revert an earlier commit

Use the type that represents the primary purpose of the commit. Do not use vague types or descriptions such as `update`, `changes`, `work`, `final`, or `misc`.

### Examples

```text
fix(butler): normalize leading zeros in numeric SKUs
feat(foa): include vendor and date in export filenames
docs: explain inventory quantity thresholds
test(butler): cover numeric SKU normalization
build: exclude local virtual environments
```

### Breaking changes

Mark an intentional breaking change with `!` and explain it in the commit body or footer:

```text
feat(api)!: change inventory response format

BREAKING CHANGE: clients must read quantity from the stock field.
```

## Commit structure

- Keep each commit focused on one coherent outcome.
- Split unrelated changes into separate commits when it improves review or rollback.
- Keep tightly related implementation, tests, and documentation together when they form one complete change.
- Add a commit body only when the reason, tradeoff, migration, or non-obvious behavior needs explanation.
- Reference an issue in a footer when relevant, for example `Closes #123`.

## Push safely

Before pushing:

1. Confirm the commit contains only intended files.
2. Fetch or otherwise verify the remote state when it may have changed.
3. If the branch is behind or has diverged, do not overwrite remote work. Integrate according to repository policy or ask the user when the correct strategy is unclear.
4. Push the intended branch to its configured upstream without force.
5. Verify that the local branch and remote branch are synchronized.

After pushing, report:

- the branch pushed;
- the short commit hash and commit message;
- whether local and remote are synchronized;
- any intentionally excluded or remaining local changes.

## Recommended agent request

When this file is copied into another project, the user can say:

```text
Read git_convention_skill.md, then review, commit, and push the current changes following that convention.
```
