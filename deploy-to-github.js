const fs = require('fs-extra');
const path = require('path');
const { execSync } = require('child_process');

// Build the project
console.log('Building the project...');
execSync('npm run build', { stdio: 'inherit' });

// Create docs directory if it doesn't exist
const docsDir = path.resolve(__dirname, 'docs');
if (!fs.existsSync(docsDir)) {
  fs.mkdirSync(docsDir);
}

// Copy dist contents to docs
console.log('Copying build files to docs folder...');
fs.copySync(path.resolve(__dirname, 'dist'), docsDir);

// Copy production manifest to docs folder
console.log('Copying production manifest to docs folder...');
fs.copyFileSync(
  path.resolve(__dirname, 'manifest.prod.xml'), 
  path.resolve(docsDir, 'manifest.xml')
);

// Copy simplified taskpane HTML file to docs folder
console.log('Copying simplified taskpane HTML file to docs folder...');
fs.copyFileSync(
  path.resolve(__dirname, 'src/taskpane/taskpane.simplified.html'), 
  path.resolve(docsDir, 'taskpane.simplified.html')
);

// Copy simplified manifest to docs folder
console.log('Copying simplified manifest to docs folder...');
fs.copyFileSync(
  path.resolve(__dirname, 'manifest.simplified.xml'), 
  path.resolve(docsDir, 'manifest.simplified.xml')
);

// Create a .nojekyll file to prevent GitHub Pages from ignoring files that begin with an underscore
fs.writeFileSync(path.resolve(docsDir, '.nojekyll'), '');

console.log('Deployment files prepared successfully!');
console.log('Now commit and push the docs folder to your GitHub repository.');
