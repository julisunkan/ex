#!/bin/bash
# Start backend API and frontend dev server together
node server/index.js &
BACKEND_PID=$!

pnpm run dev &
FRONTEND_PID=$!

# On exit, kill both
trap "kill $BACKEND_PID $FRONTEND_PID 2>/dev/null" EXIT

wait
