#!/bin/bash

docker image prune -f
sleep 3
docker build -t ddv-mai-processor .
sleep 3
echo
aws ecr get-login-password --region us-east-1 | docker login --username AWS --password-stdin 026740109516.dkr.ecr.us-east-1.amazonaws.com/ddv-mai-regisrty
echo
docker tag ddv-mai-processor:latest 026740109516.dkr.ecr.us-east-1.amazonaws.com/ddv-mai-regisrty:ddv-mai-processor
echo
docker push 026740109516.dkr.ecr.us-east-1.amazonaws.com/ddv-mai-regisrty:ddv-mai-processor
