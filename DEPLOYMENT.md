# Cloud Deployment Guide

## Docker Deployment

### Quick Start

1. **Build the image:**
   ```bash
   docker build -t survey-app:latest .
   ```

2. **Run the container:**
   ```bash
   docker run -d \
     -p 8000:8000 \
     -v $(pwd)/data:/app/data \
     --name survey-app \
     --restart unless-stopped \
     survey-app:latest
   ```

3. **Or use docker-compose:**
   ```bash
   docker-compose up -d
   ```

### Access the Application

- **Survey Form**: http://localhost:8000/
- **Admin Interface**: http://localhost:8000/admin
- **Health Check**: http://localhost:8000/health
- **API Docs**: http://localhost:8000/docs

## Cloud Deployment Options

### Option 1: Docker on Cloud VM

1. Copy project to cloud server
2. Build and run Docker container
3. Configure reverse proxy (nginx) if needed
4. Set up SSL/TLS certificates

### Option 2: Container Registry (Docker Hub, AWS ECR, etc.)

1. **Tag and push image:**
   ```bash
   docker tag survey-app:latest your-registry/survey-app:latest
   docker push your-registry/survey-app:latest
   ```

2. **Pull and run on cloud:**
   ```bash
   docker pull your-registry/survey-app:latest
   docker run -d -p 8000:8000 -v ./data:/app/data survey-app:latest
   ```

### Option 3: Cloud Platforms

#### AWS (ECS/Fargate)
- Use ECS task definition
- Mount EFS for persistent storage
- Configure ALB for load balancing

#### Google Cloud (Cloud Run)
- Deploy as Cloud Run service
- Use Cloud Storage for database backup

#### Azure (Container Instances)
- Deploy as Azure Container Instance
- Use Azure Files for persistent storage

## Security Considerations

✅ **Implemented:**
- Non-root user in container
- Health checks
- Environment variables for configuration
- Volume mounting for data persistence

⚠️ **Recommended for production:**
- Add authentication/authorization
- Use HTTPS (SSL/TLS)
- Configure firewall rules
- Set up database backups
- Use secrets management for sensitive data
- Enable rate limiting
- Add CORS configuration if needed

## Environment Variables

Create a `.env` file or set environment variables:

```bash
DATABASE_URL=sqlite:///./data/research_data.db
DB_DIR=./data
```

## Data Persistence

The database is stored in the `./data` directory, which is mounted as a volume.
This ensures data persists even when the container is restarted.

## Monitoring

- Health check endpoint: `/health`
- Logs: `docker logs survey-app`
- Stats: `docker stats survey-app`

