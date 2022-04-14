# base image
FROM python:3.6.8-slim

RUN pip install flask
RUN pip install openpyxl
RUN pip install requests

WORKDIR /usr/src/app

COPY ./page.py ./
COPY ./start.sh ./

RUN chmod 777 ./start.sh

CMD [ "sh", "./start.sh" ]