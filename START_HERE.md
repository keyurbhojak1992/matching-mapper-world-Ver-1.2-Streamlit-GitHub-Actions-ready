# Start Here — 7 steps

This project is prepared for:

```text
Private Streamlit Community Cloud app
        ↓
Private GitHub draft release (temporary input files)
        ↓
GitHub Actions 8 GB worker
        ↓
Batches of salespeople
        ↓
Private combined result ZIP
```

## 1. Create or use a private GitHub repository

Upload **the contents of this folder** to the root of your private repository. The repository must contain `app.py` and `.github/workflows/process-suggested-cloud.yml` at the top level.

## 2. Enable GitHub Actions write access

In the repository, open:

```text
Settings → Actions → General
```

Allow GitHub Actions, then under **Workflow permissions** select:

```text
Read and write permissions
```

Save.

## 3. Create one fine-grained GitHub token

Open GitHub:

```text
Profile picture → Settings → Developer settings
→ Personal access tokens → Fine-grained tokens
→ Generate new token
```

Restrict it to this one private repository and grant:

```text
Actions: Read and write
Contents: Read and write
```

Copy the token. Never commit it to the repository.

## 4. Deploy the Streamlit app

Open Streamlit Community Cloud and create an app with:

```text
Repository: your private repository
Branch: main
Main file: app.py
Python: 3.12
```

In **Advanced settings → Secrets**, paste:

```toml
[github]
repository = "YOUR_USERNAME/YOUR_PRIVATE_REPOSITORY"
token = "YOUR_FINE_GRAINED_GITHUB_TOKEN"
workflow = "process-suggested-cloud.yml"
ref = "main"
```

Deploy.

## 5. Keep it private and invite colleagues

In the deployed app, open **Share** and choose:

```text
Only specific people can view this app
```

Invite trusted colleagues by email.

## 6. First test

In **Suggested Stock Mapper**:

```text
Processing destination: GitHub Actions full run
Batch scope: Only salespeople selected above
Salespeople: 1
Batch size: 1
Continue after failed batch: ON
```

Upload the real input files, submit, and use **Refresh status**. When ready, create the result link and download the ZIP.

## 7. Full run

After the one-person test succeeds:

```text
Batch scope: All salespeople discovered from current inputs
Batch size: 1 first; then try 2 or 3
Continue after failed batch: ON
```

Download the result, then tick **Allow cleanup** and click **Delete job files**.

For screenshots, troubleshooting, security details, and limits, read `STREAMLIT_GITHUB_SETUP.md`.
