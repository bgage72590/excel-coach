#!/usr/bin/env bash
# Publishes the app's code to the public repository that hosts GitHub Pages.
#
# The public repository gets a separate history: each publish is one commit on the local `public`
# branch, made from HEAD minus the research and planning notes, under your GitHub no-reply
# address. Your own history (with its author email) never leaves this machine.
#
#   scripts/publish-public.sh            commit HEAD's app files to `public` and push it as main
#
# Pushing to main starts the Pages workflow, which tests, builds and deploys the site.
set -euo pipefail
cd "$(dirname "$0")/.."

REMOTE="${PUBLIC_REMOTE:-public}"
NAME="${PUBLIC_NAME:-bgage72590}"
EMAIL="${PUBLIC_EMAIL:-269249560+bgage72590@users.noreply.github.com}"
# Kept out of the public copy.
PRIVATE=(research SCOPE.md V2-PLAN.md TESTING.md)

if [ -n "$(git status --porcelain)" ]; then
  echo "Commit or stash your changes first; this publishes HEAD." >&2
  exit 1
fi

index="$(mktemp)"
trap 'rm -f "$index"' EXIT
GIT_INDEX_FILE="$index" git read-tree HEAD
GIT_INDEX_FILE="$index" git rm --cached -r -q --ignore-unmatch "${PRIVATE[@]}"
tree="$(GIT_INDEX_FILE="$index" git write-tree)"

parent="$(git rev-parse -q --verify refs/heads/public || true)"
if [ -n "$parent" ] && [ "$(git rev-parse "$parent^{tree}")" = "$tree" ]; then
  echo "The public copy is already up to date."
else
  message="Update Excel Coach ($(date +%Y-%m-%d))

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
  commit="$(GIT_AUTHOR_NAME="$NAME" GIT_AUTHOR_EMAIL="$EMAIL" GIT_COMMITTER_NAME="$NAME" GIT_COMMITTER_EMAIL="$EMAIL" \
    git commit-tree "$tree" ${parent:+-p "$parent"} -m "$message")"
  git update-ref refs/heads/public "$commit"
  echo "Committed $(git rev-parse --short "$commit") to the public branch."
fi

git push "$REMOTE" public:main
