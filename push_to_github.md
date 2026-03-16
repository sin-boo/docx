# Push this repo to GitHub as "reusre-maker"

1. **Create the repo on GitHub**
   - Go to https://github.com/new
   - Repository name: **reusre-maker**
   - Leave "Add a README" unchecked (this folder already has one)
   - Click **Create repository**

2. **Add remote and push** (replace `YOUR_USERNAME` with your GitHub username):

   ```powershell
   cd d:\docx
   git remote add origin https://github.com/YOUR_USERNAME/reusre-maker.git
   git branch -M main
   git push -u origin main
   ```

   If your default branch is already `master`, use:

   ```powershell
   git remote add origin https://github.com/YOUR_USERNAME/reusre-maker.git
   git push -u origin master
   ```

   GitHub may show you the exact URL and commands after you create the repo—you can copy from there.
