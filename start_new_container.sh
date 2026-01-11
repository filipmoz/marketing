#!/bin/sh
sudo  docker run -d -p 8000:8000 -v /Users/filipfm/projects/marketing/data:/app/data --name survey-app survey-app:latest


