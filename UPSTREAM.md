# Upstream Policy

This directory is a nested git repository for `m365-copilot-bun-proxy`.

It tracks the upstream project at:

```text
https://github.com/edlaver/m365-copilot-bun-proxy.git
```

`origin` is the owned canonical repository for this Codex M365 workflow. Do
not merge or rebase research repositories into this checkout.

## Local Policy

- Keep owned commits small and reviewable.
- Do not discard local commits or uncommitted changes during upstream sync.
- Do not commit raw tokens, browser state, cookies, HAR files, account identifiers, or debug traces.
- Expose only `gpt-5.6-sol` at `high`; route it to M365
  `Gpt_5_6_Reasoning`.

## Safe Sync Workflow

Use the helper from the companion root:

```bash
tools/sync-proxy-upstream.sh
```

The helper fetches upstream metadata and reports divergence. It refuses to modify history when the proxy worktree is dirty.

If the owned canonical branch has new commits and the worktree is clean, review
them first:

```bash
git -C m365-copilot-bun-proxy log --oneline --decorate --left-right --cherry-pick main...origin/main
```

Then choose a normal git operation deliberately:

```bash
git -C m365-copilot-bun-proxy rebase origin/main
```

## Required Verification

After any sync, local overlay edit, dependency change, or captured M365 API refresh, run:

```bash
git -C m365-copilot-bun-proxy diff --check
git -C m365-copilot-bun-proxy status --short --branch
cd m365-copilot-bun-proxy && bun test
../m365-codex-yolo --verify
```

For a full stack check from the companion root, run:

```bash
./m365-codex-yolo --doctor
```

## Merge Expectations

Upstream sync should preserve these local capabilities unless intentionally replaced:

- Codex Responses API provider compatibility.
- `gpt-5.6-sol` model alias and `Gpt_5_6_Reasoning` Substrate selector.
- local shell/file tool-call response shaping.
- headless and visible M365 token refresh flow.
- cached token metadata support for non-JWT substrate tokens.
- hardened WebSocket handling across Bun and Node-style APIs.
