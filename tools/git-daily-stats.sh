#!/bin/bash

# Function to validate date format
validate_date() {
    if date -j -f "%Y-%m-%d" "$1" >/dev/null 2>&1; then
        return 0
    else
        return 1
    fi
}

# Get the date parameter or use yesterday's date
if [ -z "$1" ]; then
    date_param=$(date -v-1d +%Y-%m-%d)
else
    if validate_date "$1"; then
        date_param=$1
    else
        echo "Error: Invalid date format. Please use YYYY-MM-DD."
        exit 1
    fi
fi

# Calculate the next day
next_day=$(date -j -v+1d -f "%Y-%m-%d" "$date_param" +%Y-%m-%d)

# Run the git command
git log --author="$(git config user.name)" --since="$date_param" --until="$next_day" --pretty=tformat: --numstat | awk '{ add += $1; subs += $2 } END { print "Date: '"$date_param"'\nAdded lines: "add"\nRemoved lines: "subs }'