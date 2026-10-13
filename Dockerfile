FROM debian:bookworm-slim@sha256:7c7b2c966bc9ee8cedfeef67e0e279108992c77681fa595db4a9d65c06ccc587

ENV DEBIAN_FRONTEND=noninteractive LANG=C.UTF-8 LC_ALL=C.UTF-8 \
    PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 SURVEY_ROOT=/work
RUN apt-get update && apt-get install -y --no-install-recommends \
    texlive-xetex texlive-latex-extra texlive-lang-chinese texlive-lang-arabic \
    fonts-noto-core fonts-noto-cjk latexmk \
    python3 python3-openpyxl python3-pypdf poppler-utils \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /work
COPY scripts /opt/survey/scripts
ENTRYPOINT ["python3", "/opt/survey/scripts/build.py"]
