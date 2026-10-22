import os
import sys
from huggingface_hub import HfApi

def upload():
    token = os.environ.get("HF_TOKEN")
    if not token and len(sys.argv) > 1:
        token = sys.argv[1]
        
    repo_id = "amateurmind/microplastic-detection"
    base_dir = os.path.dirname(os.path.dirname(__file__))
    zip_path = os.path.join(base_dir, "merged_dataset.zip")
    yaml_path = os.path.join(base_dir, "merged_dataset", "data.yaml")
    readme_path = os.path.join(base_dir, "merged_dataset", "README.md")
    
    api = HfApi(token=token)
    
    print("1. Uploading README.md (Dataset Card)...")
    api.upload_file(
        path_or_fileobj=readme_path,
        path_in_repo="README.md",
        repo_id=repo_id,
        repo_type="dataset",
    )
    
    print("2. Uploading data.yaml...")
    api.upload_file(
        path_or_fileobj=yaml_path,
        path_in_repo="data.yaml",
        repo_id=repo_id,
        repo_type="dataset",
    )
    
    print("3. Uploading merged_dataset.zip (~75 MB)...")
    api.upload_file(
        path_or_fileobj=zip_path,
        path_in_repo="merged_dataset.zip",
        repo_id=repo_id,
        repo_type="dataset",
    )
    
    print("\n[SUCCESS] All files uploaded successfully!")
    print(f"Repository: https://huggingface.co/datasets/{repo_id}")

if __name__ == "__main__":
    upload()
