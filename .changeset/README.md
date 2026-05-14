# Changesets

This folder is used by [Changesets](https://github.com/changesets/changesets) to track
user-visible changes and drive releases.

When you make a change that users will see, add a changeset:

```bash
pnpm changeset
```

Pure internal refactors do not need a changeset, or can use the `patch` bump with the
`chore:` prefix so they do not appear in the user-facing changelog.

See `CLAUDE.md` (section "Tech stack and conventions") for details.
