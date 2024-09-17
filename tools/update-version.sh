#!/bin/bash

# Function to display usage information
usage() {
    echo "Usage: $0 <major|minor|patch>"
    exit 1
}

# Check if an argument is provided
if [ $# -eq 0 ]; then
    usage
fi

# Get the current version
current_version=$(node -p "require('./package.json').version")

# Split the version into an array
IFS='.' read -ra version_parts <<< "$current_version"

# Update the version based on the argument
case $1 in
    major)
        ((version_parts[0]++))
        version_parts[1]=0
        version_parts[2]=0
        ;;
    minor)
        ((version_parts[1]++))
        version_parts[2]=0
        ;;
    patch)
        ((version_parts[2]++))
        ;;
    *)
        usage
        ;;
esac

# Join the version parts back into a string
new_version="${version_parts[0]}.${version_parts[1]}.${version_parts[2]}"

# Confirmation step
read -p "Are you sure you want to update from $current_version to $new_version? (y/n) " -n 1 -r
echo    # Move to a new line

if [[ $REPLY =~ ^[Yy]$ ]]
then
    # Update package.json with the new version
    sed -i '' "s/\"version\": \"$current_version\"/\"version\": \"$new_version\"/" package.json
    echo "Version updated from $current_version to $new_version"
else
    echo "Version update cancelled."
    exit 1
fi
