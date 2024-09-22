FROM python:3.12.5-bullseye

RUN apt-get update && apt-get install -y libgl1-mesa-glx

RUN apt-get install -y wget unzip bc libleptonica-dev

RUN apt-get install -y --reinstall make && \
    apt-get install -y g++ autoconf automake libtool pkg-config \
    libpng-dev libjpeg62-turbo-dev libtiff5-dev libicu-dev \
    libpango1.0-dev autoconf-archive

WORKDIR /app

RUN mkdir src && cd /app/src && \
    wget https://github.com/tesseract-ocr/tesseract/archive/refs/tags/5.4.1.zip && \
    unzip 5.4.1.zip && \
    cd /app/src/tesseract-5.4.1 && ./autogen.sh && ./configure && make && make install && ldconfig

RUN cd /usr/local/share/tessdata && wget https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata
RUN cd /usr/local/share/tessdata && wget https://github.com/tesseract-ocr/tessdata_best/raw/main/tha.traineddata

ENV TESSDATA_PREFIX=/usr/local/share/tessdata

WORKDIR /code

COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir --upgrade -r requirements.txt

COPY . .
CMD ["fastapi", "run", "app/main.py", "--port", "80"]