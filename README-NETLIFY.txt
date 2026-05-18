GREENACRE CLIENT DATABASE - NETLIFY VERSION

Folder:
C:\Users\PeterNicholson\OneDrive - Greencare Environmental\Documents\Greenacre-Client-App-Netlify

What this version does:
- Hosts the Greenacre Client Database on Netlify.
- Stores shared data through Netlify Functions and Netlify Blobs.
- Requires users to sign in before the app opens.
- Supports normal users and one or more admin users.
- Admin users can add extra asset data cells from Asset Info. Those new cells appear on all existing assets and all new assets.

Default first admin login:
Username: admin
Password: GreenacreAdmin123!

IMPORTANT: Change this after deployment by setting Netlify environment variables before first login:
GREENACRE_ADMIN_USERNAME
GREENACRE_ADMIN_PASSWORD

Recommended first setup:
1. Create a Netlify account.
2. Create a new Netlify site/project from this folder, preferably through a GitHub repo.
3. In Netlify, set these environment variables before first use:
   GREENACRE_ADMIN_USERNAME = your chosen admin username
   GREENACRE_ADMIN_PASSWORD = a strong admin password
4. Deploy the site.
5. Open the Netlify site URL.
6. Sign in as the first admin.
7. Click Admin in the app header.
8. Create the other user accounts, normally role = user.
9. Only admin users can create more users and add new asset data cells.

How data is saved:
- The browser does not write directly to a database.
- The browser calls /api/data.
- The Netlify Function checks the signed-in user session.
- The function saves the shared app state into Netlify Blobs.
- Data is versioned to reduce accidental overwrites if two users save at the same time.

Files that matter:
index.html
app.js
netlify.toml
package.json
netlify/functions/api.mts

Do not upload node_modules manually. Netlify installs dependencies from package.json.