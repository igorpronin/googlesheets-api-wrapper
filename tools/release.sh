#!/bin/bash

# Function to display usage information
usage() {
    echo "Usage: $0 <major|minor|patch> <commit message>"
    exit 1
}

# Check if both arguments are provided
if [ $# -lt 2 ]; then
    usage
fi

# Extract version type and commit message
version_type=$1
shift
commit_message="$*"

# Run update-version.sh
./tools/update-version.sh "$version_type"

# Check if update-version.sh was successful
if [ $? -eq 0 ]; then
    # If successful, run git-commit.sh
    ./tools/git-commit.sh "$commit_message"
    
    if [ $? -eq 0 ]; then
        echo "Version updated and changes committed successfully."
        
        # Run build-and-publish.sh as the last step
        ./tools/build-and-publish.sh
        
        if [ $? -eq 0 ]; then
            echo "Package built and published successfully."
        else
            echo "Failed to build and publish package."
            exit 1
        fi
    else
        echo "Failed to commit changes."
        exit 1
    fi
else
    echo "Version update failed. Commit and publish aborted."
    exit 1
fi