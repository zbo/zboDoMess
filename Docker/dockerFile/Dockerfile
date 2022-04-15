# base image
FROM python:3.6.8-slim

WORKDIR /usr/src/app

# using pip.conf or set in command line both works
# when you want to replace orignal pip source

COPY ./page.py ./
COPY ./start.sh ./
RUN mkdir /root/.pip
COPY ./pip.conf /root/.pip/pip.conf

RUN pip install -U pip
RUN pip config set global.index-url http://mirrors.aliyun.com/pypi/simple
RUN pip config set install.trusted-host mirrors.aliyun.com

RUN pip install flask
RUN pip install openpyxl
RUN pip install requests
RUN whoami
RUN pip install baostock

RUN chmod 777 ./start.sh

CMD [ "sh", "./start.sh" ]