from fastapi import FastAPI
from sqlalchemy import text

from app.config import settings
from app.database import engine, Base

from app.models.categoria import Categoria
from app.models.produto import Produto

from app.routers.categoria import router as categoria_router


app = FastAPI(
    title=settings.app_name,
    version=settings.app_version,
)

Base.metadata.create_all(bind=engine)

app.include_router(categoria_router)


@app.get("/")
def root():
    return {
        "application": settings.app_name,
        "version": settings.app_version,
    }


@app.get("/health")
def health():
    return {
        "status": "ok"
    }


@app.get("/database")
def database():

    with engine.connect() as connection:

        result = connection.execute(
            text("SELECT 1")
        )

        return {
            "database": result.scalar()
        }