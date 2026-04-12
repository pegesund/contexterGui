#!/usr/bin/env python3
"""Upload models and language data to Contabo S3."""

import boto3
import os
import sys

ENDPOINT = "https://eu2.contabostorage.com"
BUCKET = "spell"
ACCESS_KEY = "cd59e2c4bbbd7bd29951f126d87a096a"
SECRET_KEY = "3f28f3941d0d20aaa829ef17c50fe4e7"

BASE = os.path.dirname(os.path.abspath(__file__))
DYSLEX = os.path.dirname(BASE)

# Files to upload: (local_path, s3_key)
FILES = [
    # Shared BERT model
    (f"{DYSLEX}/contexter-repo/training-data/onnx/norbert4_base_int8.onnx", "models/bert/norbert4_base_int8.onnx"),
    (f"{DYSLEX}/contexter-repo/training-data/onnx/tokenizer.json",          "models/bert/tokenizer.json"),

    # Bokmal (nb)
    (f"{DYSLEX}/rustSpell/mtag-rs/data/fullform_bm.mfst",                   "lang/nb/fullform_bm.mfst"),
    (f"{DYSLEX}/contexter-repo/training-data/wordfreq_bm.tsv",              "lang/nb/wordfreq_bm.tsv"),
    (f"{DYSLEX}/syntaxer/grammar_rules.pl",                                  "lang/nb/grammar_rules.pl"),
    (f"{DYSLEX}/syntaxer/compound_data.pl",                                  "lang/nb/compound_data.pl"),
    (f"{DYSLEX}/syntaxer/sentence_split.pl",                                 "lang/nb/sentence_split.pl"),

    # Nynorsk (nn)
    (f"{DYSLEX}/rustSpell/mtag-rs/data/fullform_nn.mfst",                   "lang/nn/fullform_nn.mfst"),
    (f"{DYSLEX}/contexter-repo/training-data/wordfreq_nn.tsv",              "lang/nn/wordfreq_nn.tsv"),
    (f"{DYSLEX}/nynorsk/grammar_rules.pl",                                   "lang/nn/grammar_rules.pl"),
    (f"{DYSLEX}/nynorsk/compound_data.pl",                                   "lang/nn/compound_data.pl"),
    (f"{DYSLEX}/nynorsk/sentence_split.pl",                                  "lang/nn/sentence_split.pl"),
]

def main():
    from botocore.config import Config
    s3 = boto3.client(
        "s3",
        endpoint_url=ENDPOINT,
        aws_access_key_id=ACCESS_KEY,
        aws_secret_access_key=SECRET_KEY,
        region_name="eu2",
        config=Config(s3={"addressing_style": "path"}),
    )

    for local_path, s3_key in FILES:
        if not os.path.exists(local_path):
            print(f"SKIP (not found): {local_path}")
            continue
        size_mb = os.path.getsize(local_path) / (1024 * 1024)
        print(f"Uploading {s3_key} ({size_mb:.1f} MB) ...", end=" ", flush=True)
        s3.upload_file(local_path, BUCKET, s3_key)
        print("OK")

    # List what's in the bucket
    print("\nBucket contents:")
    resp = s3.list_objects_v2(Bucket=BUCKET)
    for obj in resp.get("Contents", []):
        size_mb = obj["Size"] / (1024 * 1024)
        print(f"  {obj['Key']:50s}  {size_mb:8.1f} MB")

if __name__ == "__main__":
    main()
