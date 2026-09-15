Office attendance import receiver for the existing DDO Node backend.

No database, no extra .env, no schema. Copy the `receiver` folder into the DDO repo and mount it on the existing Express app with the existing auth.

Mount:

```js
const { createAttendanceImportRouter } = require("./receiver");

app.use(
  "/api/attendance",
  existingAuthMiddleware,
  createAttendanceImportRouter()
);
```

That exposes:

POST /api/attendance/import

Office .env:

DDO_API_ENDPOINT=https://ddoplusnodeapi.walkingtree.tech/api/attendance/import

Need to save rows later with DDO models:

```js
createAttendanceImportRouter({
  persist: async (verified) => {
    // verified.records is the clean list
    // return { records_inserted, records_updated } if you know those counts
  },
});
```
