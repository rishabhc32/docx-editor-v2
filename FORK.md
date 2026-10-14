# Maintain this fork

This fork keeps a small set of patches on top of EigenPal's [`docx-editor`](https://github.com/eigenpal/docx-editor). Each patch is one commit with its own test, so you can move, check, and drop it alone.

## Branches

- `main` mirrors `eigenpal/docx-editor` `main`. It never holds a patch.
- `patched` is an upstream commit plus the patches, one commit for each patch.

To list the patches, run:

```bash
git log --oneline main..patched
```

## Set up a clone

Run these commands once in each clone:

```bash
git remote add upstream https://github.com/eigenpal/docx-editor.git
# Fetch only: nothing can be pushed to EigenPal by mistake.
git remote set-url --push upstream DISABLED
# Reuse the conflict fixes that you made in earlier rebases.
git config rerere.enabled true
```

## Move the patches to a new upstream release

These steps move the published patches. For the full procedure, which also covers local feature branches and a check of each patch against the release, see [Upgrade the fork to a new upstream release](FORK-UPGRADE.md).

1. Fetch upstream, and update the mirror:

   ```bash
   git fetch upstream --tags
   git push origin upstream/main:main
   ```

1. Rebase the patches onto the release tag, such as `v2.31.0`:

   ```bash
   git switch patched
   git rebase v2.31.0
   ```

   If a conflict stops the rebase, fix the files, run `git add` on them, and run `git rebase --continue`. If upstream already contains a patch, the rebase drops that commit.

1. Compare each patch before and after the rebase:

   ```bash
   git range-diff patched@{1}...patched
   ```

1. Run the checks:

   ```bash
   bun install
   bun run typecheck
   bun run lint
   bun run test
   bun run build:packages
   bun run api:check
   ```

1. Push the branch:

   ```bash
   git push --force-with-lease origin patched
   ```

## Drop a patch that upstream fixed

Upstream can fix the same bug in a different way. To check a patch, remove its source change, keep its tests, and run them. If the tests still pass, upstream covers the change.

To drop the patch, replace `<commit>` with its hash:

```bash
git rebase --onto <commit>^ <commit> patched
```

## Releases

GitHub keeps the upstream workflows off in this fork. If you turn them on in the **Actions** tab, keep the `Release` workflow off. On a push to `main`, it versions and publishes the packages with changesets.
