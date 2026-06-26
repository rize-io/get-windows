#Updating Active-Win

1. Clone the repo: [https://github.com/rize-io/active-win]
2. `cd` into the repo
3. Add the sindresorhus remote `git remote add sindresorhus https://github.com/sindresorhus/active-win.git`
4. If you run `git remote -v` you should be able to see both origin and sindresorhus remotes.
5. Checkout the `sindresorhus-main` branch. This branch is an exact copy of the original repo.
6. Run `git pull --rebase sindresorhus main` . This pulls in any recent changes from the original repo to make sure our fork stays in sync with it. Note, this is saying access the `sindresorhus` user’s repo of active win and rebase the `main` branch with Rize’s `sindresorhus-main` branch.
7. There shouldn’t be any conflicts since it’s an exact copy.
8. Go back to the main branch `git checkout main`.
9. Make any changes to any of the source files `main.cc` for Windows and `main.swift` for macOS.
10. On Windows, build the binaries from source `npm run build:windows:install -- --build-from-source --verbose`.
11. On macOS, build the binaries from source `npm run build:macos -- --verbose`.
12. On Windows, run your build `node lib/windows.js`, you can add `console.log(addon.getActiveWindow());` to the end of the file to see the output.
13. On macOS, run your build `node lib/macos.js`, you can run the `getOpenWindowsSync` function and log the results.
14. If everything looks good, commit your changes.
15. Update the npm version `npm version prerelease`.
16. Push your changes to the `main` branch.
17. Create a release on GitHub with the same prerelease version and set it to the latest release. You might have to uncheck the prerelease option.
18. This will trigger a GitHub Action that will build Node 18 binaries and attach them to the release. You see the following assets:
    1. napi-3-win32-unknown-ia32.tar.gz
    2. napi-3-win32-unknown-x64.tar.gz
    3. napi-6-win32-unknown-ia32.tar.gz
    4. napi-6-win32-unknown-x64.tar.gz
19. I'm not entirely sure why the original library isn't also uploading the macOS binaries.
20. Run `npm publish` to publish the package to npm.
