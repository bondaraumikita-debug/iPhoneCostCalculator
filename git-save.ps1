$msg = "Update: $(Get-Date -Format yyyy-MM-dd_HH-mm-ss)"
git add .
git commit -m "$msg"
git push
