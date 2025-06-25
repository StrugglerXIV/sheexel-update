#!/bin/bash

# Get the latest tag
latest_tag=$(git describe --tags --abbrev=0 2>/dev/null)

# If no tag yet, start at v0.0.0
if [ -z "$latest_tag" ]; then
  major=0
  minor=0
  patch=0
else
  IFS='.' read -r major minor patch <<< "${latest_tag#v}"
fi

echo "🔢 Current version: v$major.$minor.$patch"
read -p "👉 Enter new patch number (current is $patch): " new_patch

# Fallback to same patch if input is empty
new_patch=${new_patch:-$patch}

new_tag="v$major.$minor.$new_patch"

echo "🏷️ Tagging new version: $new_tag"
git tag "$new_tag"
git push origin "$new_tag"
