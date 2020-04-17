FROM ubuntu:18.04
MAINTAINER Sankar Biswas

RUN apt-get update \
  && apt-get install -y vim zip unzip wget curl \
  && apt-get install -y python3-pip python3-dev \
  && cd /usr/local/bin \
  && ln -s /usr/bin/python3 python \
  && pip3 install --upgrade pip

COPY requirements.txt /opt/ddv-mai/
COPY properties.ini /opt/ddv-mai/
COPY python-s3.py /opt/ddv-mai/
COPY run.sh /opt/ddv-mai/

RUN mkdir -p /opt/File/ && \
    mkdir -p /opt/File/output && \
    mkdir -p /opt/File/temp_storage && \
    mkdir -p /opt/File/intermediate_files && \
    mkdir -p /opt/File/logs && \
    touch /opt/File/logs/example.log && \
    pip install --no-cache-dir -r  /opt/ddv-mai/requirements.txt && \
    chmod +x /opt/ddv-mai/run.sh

WORKDIR /opt/ddv-mai/

CMD ["./run.sh"]
