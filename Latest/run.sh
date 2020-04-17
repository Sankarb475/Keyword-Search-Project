#!/bin/bash
set -x

python3 python-s3.py "${s3_vdr_folder_location}" "${s3_drl_folder_location}" "${additional_keywords}" "${guid}"
#python3 python-s3.py "s3://aws-workdocs-test-bucket/sankar.biswas@pwc.com/VDR" "s3://aws-workdocs-test-bucket/sankar.biswas@pwc.com/DRL" "Tum, ho" "sankar.biswas@pwc.com"
