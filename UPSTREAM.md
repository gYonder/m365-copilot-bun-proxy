# Upstream Policy

This directory is a nested git repository for `m365-copilot-bun-proxy`.

It tracks the upstream project at:

```text
https://github.com/edlaver/m365-copilot-bun-proxy.git
```

This local checkout is not expected to push changes upstream. Local commits are an intentional overlay for the Companion M365 Copilot + Codex workflow.

## Local Policy

- Keep upstream as a read-only source of proxy updates.
- Keep local overlay commits small and reviewable.
- Do not discard local commits or uncommitted changes during upstream sync.
- Do not commit raw tokens, browser state, cookies, HAR files, account identifiers, or debug traces.
- Treat `gpt-5.5` as the local proxy alias for M365 Copilot `Gpt_5_5_Reasoning`.

## Safe Sync Workflow

Use the helper from the companion root:

```bash
tools/sync-proxy-upstream.sh
```

The helper fetches upstream metadata and reports divergence. It refuses to modify history when the proxy worktree is dirty.

If upstream has new commits and the worktree is clean, review the commits first:

```bash
git -C m365-copilot-bun-proxy log --oneline --decorate --left-right --cherry-pick main...origin/main
```

Then choose a normal git operation deliberately, usually rebase for a linear local overlay:

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
- `gpt-5.5` model alias and GPT-5.5 Substrate tone mapping.
- local shell/file tool-call response shaping.
- headless and visible M365 token refresh flow.
- cached token metadata support for non-JWT substrate tokens.
- hardened WebSocket handling across Bun and Node-style APIs.
