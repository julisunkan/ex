#!/bin/bash
# Kill any leftover server on port 3001
fuser -k 3001/tcp 2>/dev/null || true

# Start backend API
node server/index.js &
BACKEND_PID=$!

# Start frontend dev server
pnpm run dev -- --config vite.config.ts &
FRONTEND_PID=$!

# On exit, kill both
trap "kill $BACKEND_PID $FRONTEND_PID 2>/dev/null" EXIT

wait
