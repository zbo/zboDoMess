# base image
FROM python:3.6.8-slim

WORKDIR /usr/src/app

COPY ./page.py ./
COPY ./start.sh ./
RUN mkdir /root/.pip
COPY ./pip.conf /root/.pip/pip.conf

RUN pip install flask
RUN pip install openpyxl
RUN pip install requests
RUN whoami
RUN pip install baostock

RUN chmod 777 ./start.sh

CMD [ "sh", "./start.sh" ]