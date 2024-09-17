const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const watch_dir = './src'; // Directory to watch
const command_to_run = 'npm run clear && npm run build';

console.log(`Watching for file changes in ${watch_dir}`);

fs.watch(watch_dir, { recursive: true }, (eventType, filename) => {
  if (filename) {
    const filePath = path.join(watch_dir, filename);
    console.log(`File ${filePath} has been ${eventType}`);

    exec(command_to_run, (error, stdout, stderr) => {
      if (error) {
        console.error(`Error executing command: ${error}`);
        return;
      }
      if (stdout) console.log(`Command output: ${stdout}`);
      if (stderr) console.error(`Command errors: ${stderr}`);
    });
  }
});

// Prevent the Node.js process from exiting
process.stdin.resume();

console.log('File watcher started. Press Ctrl+C to exit.');
