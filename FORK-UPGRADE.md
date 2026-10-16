# Upgrade the fork to a new upstream release

When EigenPal publishes a release, move all fork work onto it: the mirror, the published patches, and the local feature branches that are not published yet. Then check every patch and feature against the release, because the release can make a patch unnecessary, offer a better way to write it, or break it.

[Maintain this fork](FORK.md) describes the branches and the basic rebase. This guide is the full procedure. In the examples, `v2.30.0` is the old release and `v2.31.0` is the new one. Replace them with your versions.

## Know the branches

| Branch | Holds | Pushed |
| --- | --- | --- |
| `main` | A mirror of `eigenpal/docx-editor` `main`. Never a patch. | Yes |
| `patched` | An upstream release plus the published patches, one commit for each patch. | Yes |
| Feature branches, such as `fix/<name>` | One feature that is not in `patched` yet. One or more commits. | No |
| `candidate/prod-v<release>` | `patched` plus every feature commit, in order. The integration branch that you test before a release. | No |
| `prod-v<release>` | A release tag plus the patches. The source of the tarball releases. | Yes |

A feature branch sits on `main`, or on the branch it needs. For example, a branch that extends tracked tables sits on `fix/tracked-table-edits`.

## Before you start

1. Check the clone setup in [Maintain this fork](FORK.md#set-up-a-clone), including `git config rerere.enabled true`.
1. Fetch the release:

   ```bash
   git fetch upstream --tags
   git log --oneline -1 v2.31.0 upstream/main
   ```

1. List the upstream changes. You read them again in the evaluation steps:

   ```bash
   git log --oneline v2.30.0..v2.31.0
   ```

1. Confirm that GitHub has not registered the upstream workflows in the fork. A push to `main` runs the upstream `Release` workflow if it is registered.

   ```bash
   gh api repos/<owner>/docx-editor-v2/actions/workflows --jq '.workflows[].name'
   ```

   If the list includes `Release`, turn it off in the **Actions** tab before you push.

## Step 1: Move `main` to the release

```bash
git branch -f main v2.31.0
git push origin main:main
```

The push is a fast-forward. After it, the commit hook's collaboration check and every diff against `main` use the new base.

## Step 2: Rebase the published patches

Work in a separate worktree, so that you can compare the result with the old branch:

```bash
git worktree add -b eval/patched-v2.31.0 ../wt/patched-eval patched
cd ../wt/patched-eval
git rebase v2.31.0
```

For each conflict, understand both sides. Prefer the upstream structure, and express the patch's intent on top of it. Resolve API snapshot conflicts (`docs/api/**`) as well as you can, then regenerate them in step 4.

`rerere` reuses a fix that you recorded earlier for the same conflict. It does not check whether that fix is still right. Review every file that `rerere` resolves.

## Step 3: Evaluate each published patch

For each patch, answer these questions and record the answers:

1. **Does upstream cover it now?** Remove the patch's source change, keep its tests, and run them. If the tests pass, upstream covers the change. Drop the patch as [Maintain this fork](FORK.md#drop-a-patch-that-upstream-fixed) describes.
1. **What did upstream change nearby?** Read the upstream changes in the files that the patch touches:

   ```bash
   git log -p v2.30.0..v2.31.0 -- <files the patch touches>
   ```

1. **Can the patch be smaller?** Look for a new upstream helper, seam, or pattern that replaces code in the patch. Prefer fewer changed lines in EigenPal files.
1. **Does a new upstream path need the patch's behaviour?** A new command, story, or lane can get around a patch. For example, upstream bookmark commands could change a header while an option made headers read-only.
1. **Does a new upstream feature break the patch?** For example, a release that changed where typing lands beside a bookmark broke a patch that inserted text at a link's end.

To change a patch, keep it one commit:

```bash
git commit --fixup <patch commit>
HUSKY=0 GIT_SEQUENCE_EDITOR=true git rebase -i --autosquash v2.31.0
```

Each patch keeps its own test, and that test must fail without the patch's source change.

## Step 4: Check the published patches

Build first. The `pro` type check reads the built `i18n` and `core` packages, so a stale build reports false errors.

```bash
bun install
bun run build:packages
node scripts/collaboration/cli.mjs check --base v2.31.0
bun run typecheck
bun run lint
bun run format:check
bun run api:check
bun run i18n:validate
```

If `api:check` reports drift, run `bun run api:extract`, and put the snapshot change in the commit that caused it. Run every test file that each patch adds or changes. Then compare each patch before and after the rebase:

```bash
git range-diff v2.30.0..patched v2.31.0..eval/patched-v2.31.0
```

## Step 5: Rebase the local feature branches

List the feature branches and their commits:

```bash
for b in $(git branch --format='%(refname:short)' | grep -E '^(fix|feat)/'); do
  echo "== $b"; git log --oneline "$(git merge-base "$b" v2.31.0)..$b"
done
```

Rebase each branch in its own worktree. Rebase a branch that sits on `main` onto the release, and a stacked branch onto its rebased base:

```bash
git -C <worktree> rebase --onto v2.31.0 <old base>
git -C <stacked worktree> rebase --onto <rebased base branch> <old base tip>
```

Check each branch alone: build `i18n`, run the type checks, run lint, and run every test file that the branch adds.

## Step 6: Evaluate each feature branch

Ask the same questions as in step 3 for each feature commit. Also look for these problems:

- **Paths that get around the feature.** A new upstream lane can write directly, so that a guard such as `untrackedEdits: 'refuse'` never sees the write.
- **Features that conflict with each other.** Two fork features can each pass their own tests and fail together. For example, a guard that refuses every edit it cannot classify also refused another feature's new tracked edit. Test the features together in the candidate.
- **Fixes that are valid only in combination.** Some fixes are correct only when two patches meet, for example a test expectation that holds only after another patch tracks a change. Put such a fix in the later commit of the candidate, not in the standalone branch.

## Step 7: Build and test the candidate

```bash
git worktree add -b candidate/prod-v2.31.0 ../wt/candidate eval/patched-v2.31.0
cd ../wt/candidate
git cherry-pick <feature commits, oldest first>
```

Then run the checks from step 4 and the full suite. On a laptop, run the suite at a low priority with 3 workers:

```bash
nice -n 15 node scripts/test/run-parallel.mjs --jobs 3
```

Watch for these combination problems:

- **Line caps.** Several commits can each stay under a file's cap and together exceed it. Cut the lines in the commits that added them. Do not raise a cap.
- **Pinned lists.** Tests such as `part-for-call-sites.test.ts` and `revision-capable-ops.test.ts` pin call sites and op lists. A new feature must update them, or use a different approach.

Finally, run the app with two people in one collaboration room, and test each feature by hand: suggest, accept, reject, undo, save, and reopen.

## Step 8: Publish the branches

1. Move `patched` to the evaluated patches, and push it. The lease refuses the push if someone changed `origin/patched` in the meantime:

   ```bash
   git branch -f patched eval/patched-v2.31.0
   git push --force-with-lease=patched:<old origin/patched commit> origin patched
   ```

1. After the candidate passes, add the feature commits to `patched` in the same way.
1. Create `prod-v2.31.0` from the release tag plus the patches, and push it. If `patched` sits on the release tag, both branches are the same commit:

   ```bash
   git branch prod-v2.31.0 patched
   git push origin prod-v2.31.0
   ```

## Step 9: Build the tarballs

A fork release is a GitHub release with five npm tarballs: `core`, `react`, `pro`, `editor-api`, and `i18n`. Its version is the upstream version plus `-sp.<n>`, such as `2.31.0-sp.1`. The fork's releases are immutable, so a new build needs a new tag, such as `prod-v2.31.0-sp.2`.

1. Make a clean checkout of the release branch, and install:

   ```bash
   git worktree add --detach /tmp/fork-release/prod prod-v2.31.0
   cd /tmp/fork-release/prod
   bun install --frozen-lockfile
   ```

1. Stamp the version. Do this before the build, because `react` embeds its version:

   ```bash
   node scripts/fork/stamp-release.mjs 2.31.0-sp.1
   ```

   The script pins the links between the five packages to that version. The optional `vue` peer of `pro` keeps its upstream range. Upstream manifest tests expect upstream ranges, so they fail on a stamped tree. Run the tests on the candidate, not here.

1. Build and check the packages:

   ```bash
   bun run build:packages
   bun run notices:generate
   bun run check:package-artifacts
   ```

1. Pack the tarballs, and record their digests. `npm pack --json` also gives each tarball's `integrity` for the release notes:

   ```bash
   mkdir -p /tmp/fork-release/out
   for d in core react pro editor-api i18n; do
     (cd packages/$d && npm pack --json --ignore-scripts --pack-destination /tmp/fork-release/out > /tmp/fork-release/out/$d.json)
   done
   cd /tmp/fork-release/out && shasum -a 256 *.tgz > local-digests.txt
   ```

## Step 10: Audit the release for remote data

The fork's builds must not send telemetry or other data to a remote server by default. Audit each release before you publish it.

1. Unpack the tarballs, and list the network calls and remote hosts in the shipped code:

   ```bash
   mkdir -p /tmp/fork-release/x && cd /tmp/fork-release/x
   for f in ../out/*.tgz; do n=$(basename "$f" .tgz); mkdir -p "$n" && tar -xzf "$f" -C "$n"; done
   J=(--include='*.js' --include='*.mjs' --include='*.cjs')
   grep -rlE '\bfetch\(|sendBeacon|XMLHttpRequest|new WebSocket|EventSource\(|importScripts|RTCPeerConnection|node:https?|node:net|node:dgram|node:dns' "${J[@]}" .
   grep -rohE '(https?|wss?)://[A-Za-z0-9.-]+' "${J[@]}" . | sort | uniq -c | sort -rn
   grep -rlE '@docx-editor\.dev/telemetry|opentelemetry' "${J[@]}" .
   ```

1. Read each result in its context. A network call must run only on a host option or a user action. These items were known in `2.31.0`:
   - **Help** > **Report issue** in `<DocxEditor>` opens github.com with page details. A host turns it off with `menu={{ reportIssue: false }}`.
   - WebRTC collaboration in `pro` needs a host's `signaling` option. Without `iceServers`, it uses the public STUN servers of simple-peer.
   - `react` fetches only a URL that the host gives.
   - `core` loads its HarfBuzz WebAssembly file from the page's own origin.

1. Compare the runtime `dependencies` with the previous release. Ignore the version changes of `@docx-editor.dev` packages. Audit each new dependency in the same way:

   ```bash
   for d in core react pro editor-api i18n; do
     diff <(git show v2.30.0:packages/$d/package.json | jq .dependencies) \
       <(git show prod-v2.31.0:packages/$d/package.json | jq .dependencies) && echo "$d: unchanged"
   done
   ```

## Step 11: Test an install

Install the tarballs in a scratch project, as an app does. Put each tarball's `file:` path in both `dependencies` and `overrides`, so each link between the packages resolves to this build. Add `ignore-scripts=true` to the project's `.npmrc`.

- **Server.** Install `core`, `pro`, and `editor-api`. Open a small document with `DocxEditor.createServer`, set `changeTrackingMode`, and replace a word with `insertText(text, 'Replace')`. The saved `word/document.xml` must have a `w:ins` and a `w:del`.
- **Browser.** Install `core`, `react`, and `pro` with React. Import `@docx-editor.dev/react`, `@docx-editor.dev/pro`, and `@docx-editor.dev/core/editor`. Each import must load.

## Step 12: Create and publish the release

1. Write the release notes in `/tmp/fork-release/notes.md`:
   - The source branch and commit, and the upstream tag.
   - The version, and the build commands from step 9.
   - A table of the five packages with their tarball URLs and `integrity` values.
   - An `overrides` block for the five tarball URLs. A version such as `2.31.0-sp.1` sorts before `2.31.0`, so a range such as `^2.31.0` does not match it. Apps must use the URLs.
   - If the collaboration format changed, say that saved rooms must be created again.
   - The patch list:

     ```bash
     git log --reverse --format='- %s' v2.31.0..prod-v2.31.0
     ```

1. Create a draft release. With `--target`, GitHub creates the tag only when you publish:

   ```bash
   SHA=$(git -C /tmp/fork-release/prod rev-parse HEAD)
   cd /tmp/fork-release/out
   gh release create prod-v2.31.0-sp.1 --repo <owner>/docx-editor-v2 --draft \
     --target "${SHA}" --title prod-v2.31.0-sp.1 --notes-file ../notes.md \
     docx-editor.dev-{core,react,pro,editor-api,i18n}-2.31.0-sp.1.tgz
   ```

1. Compare the uploaded digests with the local digests. `diff` prints nothing when they match:

   ```bash
   gh api repos/<owner>/docx-editor-v2/releases \
     --jq '.[] | select(.tag_name == "prod-v2.31.0-sp.1") | .assets[] | "\(.digest | sub("sha256:"; ""))  \(.name)"' \
     | sort | diff - <(sort local-digests.txt)
   ```

1. Publish the draft only after the fork owner approves. You cannot change a published release:

   ```bash
   gh release edit prod-v2.31.0-sp.1 --repo <owner>/docx-editor-v2 --draft=false
   ```

1. Check the published release:
   - The tag points at the release commit, and the release is immutable:

     ```bash
     git fetch origin --tags
     git rev-parse 'prod-v2.31.0-sp.1^{commit}'
     gh api repos/<owner>/docx-editor-v2/releases/tags/prod-v2.31.0-sp.1 --jq .immutable
     ```

   - Download each tarball without a login, and compare its digest with `local-digests.txt`.
   - Repeat the server test from step 11 with the public URLs.

1. Remove the build checkout. It holds only the stamped `package.json` files:

   ```bash
   git worktree remove --force /tmp/fork-release/prod
   ```

## Keep the collaboration format apart from upstream

If a fork patch changes the collaboration format, for example by raising `repairVersion` in `document-compatibility.ts`, the fork's format number can collide with a future upstream number. A collision lets two builds with different rules join one room without a refusal.

- When a rebase meets an upstream format change, give the fork a number higher than upstream's.
- Every saved room in the old format must be exported and seeded again. Plan this before you release a format change to an app that keeps saved rooms.

## Avoid known pitfalls

- **Rebases and hooks.** While `origin/main` is still the old base, the commit hook's collaboration check fails on upstream changesets. Run `HUSKY=0 git ...` for the rebase and cherry-pick commits, and run the checks by hand with `--base v2.31.0`.
- **Interactive rebase.** Use `GIT_SEQUENCE_EDITOR=true git rebase -i --autosquash <base>` for a non-interactive autosquash.
- **Shell loops.** zsh does not split an unquoted variable into words. Run loops over file lists in `bash`.
- **zsh refspecs.** zsh reads `$VAR:r` as a modifier, so `git push origin $SHA:refs/heads/patched` pushes a broken refspec. Write `"${SHA}:refs/heads/patched"`.
- **Machine load.** `bun run test` sets its width from the memory, not from the CPU load. On a laptop, a full suite with 8 workers beside package builds made the machine stop responding. Run one heavy task at a time, with `nice` and `--jobs 3`.
- **API check.** `api:check` needs every package built, including `docx-to-pdf`.
- **Old commits.** Move a ref with `git branch -f` only after you check that the old commits are on another branch or in the reflog.

## Ask Claude to do the upgrade

You can give this request to Claude Code in this repository:

> Upgrade the fork to `<release tag>` as FORK-UPGRADE.md describes. Move `main`, rebase the published patches and the local feature branches, evaluate each patch and feature against the release, fix what needs fixing, build and test the candidate, and report. Do not push until I approve.

To build a release, you can give this request:

> Build the fork release `prod-v<release>-sp.<n>` from `prod-v<release>` as FORK-UPGRADE.md describes, steps 9 to 12. Run the audit and the install test, and create a draft release. Do not publish until I approve.
