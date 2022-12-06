FROM python:latest
COPY  . .
RUN pip install -r requirements.txt 
CMD [ "python", "./crear_excel.py" ]