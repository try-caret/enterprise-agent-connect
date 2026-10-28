# Releasing

`.github/workflows/release.yml` publishes three things from one commit: the server package
`@arbium/enterprise-agent-connect`, the client `@arbium/enterprise-agent-connect-client` (both on
npmjs, with provenance) and the image `ghcr.io/try-caret/enterprise-agent-connect`. Before any of
it, it runs the whole CI on that commit: check, tests on Postgres, the consumer test
(`scripts/consumer-test.sh`: the packed packages installed in a fresh project, type-checked and
run), the image and the chart. Nothing is published if any of it fails.

The version comes from the tag or the run, never from `package.json` (which stays as is).

## Cut a release

1. Merge what goes in into `main`.
2. Tag the commit and push the tag:
   ```sh
   git tag v0.2.0 && git push origin v0.2.0
   ```
   `v0.2.0` publishes npm dist-tag `latest` and image tags `0.2.0` and `latest`. A prerelease
   (`v0.2.0-rc.1`) publishes dist-tag `next` and image tag `0.2.0-rc.1` only.
3. The run refuses a tag whose version is not greater than the newest on npm (dev builds aside),
   then creates the GitHub Release with generated notes (marked prerelease for `next`).

## Cut a dev build

To try an unmerged change in a host app: Actions, Release, "Run workflow", pick the branch.
Or `gh workflow run release.yml --ref <branch>`. It publishes
`<next patch>-dev.<short sha>` (for example `0.1.2-dev.3f9c2ab`) as npm dist-tag `dev` and the image
with the same tag, and makes no GitHub Release. Install it by its exact version:

```sh
npm install @arbium/enterprise-agent-connect@0.1.2-dev.3f9c2ab
```

## Switch to trusted publishing

Today the job authenticates with the `NPM_TOKEN` secret. It runs npm 11.5 or later, which tries
OpenID Connect first, so trusted publishing takes over as soon as it is configured:

1. On npmjs.com, for each of the two packages: Settings, Trusted publishing, add GitHub Actions
   with organization `try-caret`, repository `enterprise-agent-connect`, workflow `release.yml`.
2. Cut a dev build and check that both packages published, with provenance on their npm pages.
3. Delete the `NPM_TOKEN` repository secret and remove `NODE_AUTH_TOKEN` from the `npm` job.
   Optionally, in each package's settings, require trusted publishing and disallow tokens.
