#!/bin/bash

# Check if a commit message is provided
if [ $# -eq 0 ]; then
    echo "Error: Please provide a commit message."
    echo "Usage: $0 <commit message>"
    exit 1
fi

# Combine all arguments into a single commit message
commit_message="$*"

# Run git add .
echo "Staging changes..."
git add .

if [ $? -ne 0 ]; then
    echo "Error: Failed to stage changes."
    exit 1
fi

# Run git commit with the provided message
echo "Committing changes..."
git commit -m "$commit_message"

if [ $? -ne 0 ]; then
    echo "Error: Failed to commit changes."
    exit 1
fi

# Push changes to remote repository
echo "Pushing changes to remote repository..."
git push

if [ $? -ne 0 ]; then
    echo "Error: Failed to push changes."
    exit 1
fi

echo "Changes successfully committed and pushed with message: $commit_message"