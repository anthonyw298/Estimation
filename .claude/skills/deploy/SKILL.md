---
name: deploy
description: Deploy the web app to Vercel
disable-model-invocation: true
allowed-tools: Bash(vercel *), Bash(npm run build), Bash(git *)
---

# Deploy to Vercel

Deploy the web frontend to Vercel.

## Steps

1. Ensure the build passes first:
   ```bash
   cd web && npm run build
   ```

2. If build fails, fix the errors before deploying.

3. Deploy with Vercel CLI:
   - **Preview**: `cd web && vercel`
   - **Production**: `cd web && vercel --prod`

4. Report the deployment URL when complete.

## Arguments

Pass `prod` or `production` to deploy to production. Otherwise defaults to a preview deployment.
