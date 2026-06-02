FROM php:8.2-apache

RUN apt-get update && apt-get install -y \
    libzip-dev zip unzip \
    && docker-php-ext-install zip

RUN a2enmod rewrite

COPY . /var/www/html/

RUN mkdir -p /var/www/html/pdfs \
             /var/www/html/uploads \
    && chmod 777 /var/www/html/pdfs \
                 /var/www/html/uploads

WORKDIR /var/www/html
EXPOSE 80