#!/bin/bash
# Installiert die graphify-CLI in Claude-Code-Cloud-Sessions (Skill liegt in .claude/skills/graphify).
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

GRAPHIFY_VERSION="0.9.74"

if ! command -v graphify >/dev/null 2>&1 || ! pip show graphifyy 2>/dev/null | grep -q "^Version: ${GRAPHIFY_VERSION}$"; then
  pip install -q "graphifyy==${GRAPHIFY_VERSION}" 2>/dev/null \
    || pip install -q "graphifyy==${GRAPHIFY_VERSION}" --break-system-packages
fi
