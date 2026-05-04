import packager from 'electron-packager';

async function bundleApp() {
    console.log('Starting packaging...');
    const options = {
        dir: '.',
        name: 'VP Parcel',
        platform: 'win32',
        arch: 'x64',
        out: 'release-packager',
        overwrite: true,
        ignore: [
            /^\/src/,
            /^\/public/,
            /^\/release/,
            /^\/release-packager/,
            /.*\.zip$/,
            /.*\.log$/,
            /^\/vite\.config\.js$/,
            /^\/index\.html$/,
            /^\/\.git/,
            /^\/update_template\.js$/,
            /^\/package_app\.js$/
        ],
        prune: true
    };

    try {
        const appPaths = await packager(options);
        console.log(`Electron app bundles created at:`);
        console.log(appPaths);
    } catch (err) {
        console.error('Packaging failed:', err);
        process.exit(1);
    }
}

bundleApp();
