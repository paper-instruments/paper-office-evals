# Self-contained treatment-specific Harbor environment.
# The build context contains only the selected treatment distribution.

FROM python:3.13-slim@sha256:eb43ff125d8d58d7449dcba7d336c23bcac412f526d861db493b9994d8010280

# Headless office runtime, identical across package and skill treatments.
RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      libreoffice-calc fonts-dejavu-core fonts-liberation poppler-utils \
    && soffice --version \
    && rm -rf /var/lib/apt/lists/*

COPY wheelhouse/shared/ /tmp/xlsx-shared-wheels/
RUN python -m pip install --no-cache-dir --no-index \
      --find-links=/tmp/xlsx-shared-wheels \
      et_xmlfile==2.0.0 \
      lxml==6.1.1 \
      Pillow==12.3.0 \
    && rm -rf /tmp/xlsx-shared-wheels

# Skill helpers, identical across treatments; omit markitdown's openpyxl extra.
RUN python -m pip install --no-cache-dir --only-binary=:all: \
      pandas==2.3.3 markitdown==0.1.3 tabulate==0.9.0 \
      numpy==2.5.3 python-dateutil==2.9.0.post0 pytz==2026.3.post1 tzdata==2026.4 \
      beautifulsoup4==4.15.0 charset-normalizer==3.5.1 defusedxml==0.7.1 \
      magika==0.6.3 markdownify==1.2.3 requests==2.34.2 click==8.5.0 \
      onnxruntime==1.30.0 python-dotenv==1.2.3 flatbuffers==25.12.19 \
      packaging==26.3 protobuf==7.36.1 six==1.17.0 soupsieve==2.9.2 \
      typing-extensions==4.16.0 idna==3.19 urllib3==2.8.0 certifi==2026.7.22 \
    && printf '%s\n' '/home/agent/.config/opencode/skills/xlsx/scripts' \
      > /usr/local/lib/python3.13/site-packages/xlsx_skill_helpers.pth \
    && printf '%s\n' '# Block startup hooks in the writable skill directory.' \
      > /usr/local/lib/python3.13/site-packages/sitecustomize.py

COPY images/common/xlsx_treatment_fingerprint.py /usr/local/bin/xlsx-ablation-fingerprint
COPY images/common/treatment_entrypoint.sh /usr/local/bin/xlsx-ablation-entrypoint
COPY images/common/verify_treatment.sh /usr/local/bin/xlsx-ablation-verify-treatment

RUN useradd --create-home --shell /bin/bash agent \
    && mkdir -p /app /opt/xlsx-ablation \
    && chown agent:agent /app \
    && chmod 0755 \
      /usr/local/bin/xlsx-ablation-fingerprint \
      /usr/local/bin/xlsx-ablation-entrypoint \
      /usr/local/bin/xlsx-ablation-verify-treatment

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONPATH=/usr/local/lib/python3.13/site-packages \
    PYTHONSAFEPATH=1 \
    PYTHONNOUSERSITE=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_ROOT_USER_ACTION=ignore

WORKDIR /app
ENTRYPOINT ["/usr/local/bin/xlsx-ablation-entrypoint"]

# Selected treatment layer
COPY wheelhouse/paper/ /tmp/xlsx-treatment-wheels/
RUN python -m pip install --no-cache-dir --no-index --no-deps \
      --find-links=/tmp/xlsx-treatment-wheels paper-xlsx==0.2.1 \
    && rm -rf /tmp/xlsx-treatment-wheels

ENV XLSX_ABLATION_CONDITION=paper
RUN python3 /usr/local/bin/xlsx-ablation-fingerprint --condition paper \
      > /opt/xlsx-ablation/treatment-fingerprint.json

# Task fixture layer
WORKDIR /app
RUN mkdir -p /app/eval_fixtures/xlsx /app/evals
COPY filesystem/cached-model.xlsx /app/eval_fixtures/xlsx/cached-model.xlsx
RUN chown -R agent:agent /app
