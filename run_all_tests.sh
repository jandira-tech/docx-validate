#!/bin/bash
FILES=$(find tests -name "*.test.ts" -o -name "*.spec.ts")
for f in $FILES; do
  echo "Running $f"
  bun x vitest run $f
  if [ $? -ne 0 ]; then
    echo "FAILED: $f"
    break
  fi
done
echo "DONE"
