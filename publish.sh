#!/usr/bin/env bash
echo 'update code'
git pull

echo 'stop server'
pkill -ef guniorn

echo 'start server'
nohup gunicorn -b 0.0.0.0:8001 app:app &
