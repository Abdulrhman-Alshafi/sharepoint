const gulp = require('gulp');
const { exec } = require('child_process');

// Serve task - starts the development server
gulp.task('serve', (done) => {
  console.log('Starting development server...');
  const serve = exec('npx heft start --clean', { cwd: __dirname });
  
  serve.stdout.on('data', (data) => {
    process.stdout.write(data);
  });
  
  serve.stderr.on('data', (data) => {
    process.stderr.write(data);
  });
  
  serve.on('exit', (code) => {
    if (code !== 0) {
      done(new Error(`Serve failed with code ${code}`));
    }
  });
});

// Bundle task - creates production bundle
gulp.task('bundle', (done) => {
  console.log('Building production bundle...');
  exec('npx heft test --clean --production', (err, stdout, stderr) => {
    if (stdout) console.log(stdout);
    if (stderr) console.error(stderr);
    if (err) {
      done(err);
    } else {
      console.log('Bundle completed successfully!');
      done();
    }
  });
});

// Package solution task
gulp.task('package-solution', (done) => {
  console.log('Packaging solution...');
  exec('npx heft test --clean --production && npx heft package-solution --production', (err, stdout, stderr) => {
    if (stdout) console.log(stdout);
    if (stderr) console.error(stderr);
    if (err) {
      done(err);
    } else {
      console.log('Package created successfully!');
      done();
    }
  });
});

// Clean task
gulp.task('clean', (done) => {
  console.log('Cleaning build artifacts...');
  exec('npx heft clean', (err, stdout, stderr) => {
    if (stdout) console.log(stdout);
    if (stderr) console.error(stderr);
    if (err) {
      done(err);
    } else {
      console.log('Clean completed!');
      done();
    }
  });
});

// Build task
gulp.task('build', gulp.series('bundle', 'package-solution'));

// Default task
gulp.task('default', gulp.series('serve'));
