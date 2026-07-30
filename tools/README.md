# Dev-tenant seeder

`npm run seed:dev` builds a SharePoint site worth inspecting: groups with real
permission assignments, lists covering every field type the inspectors render, a
document library with folders and files, and a list whose permission inheritance
is broken. It is idempotent — re-running updates in place.

```sh
export SP_TENANT=contoso.sharepoint.com
export SP_USER=…  SP_PASS=…  SP_TOTP_SECRET=…
npm run seed:dev
```

Requires site-collection admin, and a **disposable** tenant: it creates and
modifies lists, groups and permissions.

## Persona accounts

Person columns (`Owner`, `Assigned To`, `Reported By`) and group membership are
assigned to the fictional principals listed in `PERSONAS` in `seed-data.js`:

| Display name | Username        |
| ------------ | --------------- |
| Dana Reyes   | `dana.reyes`    |
| Marcus Webb  | `marcus.webb`   |
| Priya Nair   | `priya.nair`    |

They only need to exist in the directory; a licence is not required, since
nothing ever signs in as them. Create them once in the Microsoft 365 admin
centre (**Users → Active users → Add a user**, choosing *Create user without
product license*), using the tenant's own domain.

The seeder cannot create accounts itself, so any persona missing from the
directory is skipped with a warning. If none resolve, everything is assigned to
the account running the seeder.

Why fictional accounts rather than whoever the tenant already contains: seeded
data ends up in screenshots and screen recordings committed to this repo, and
real colleagues' names should not.
