# Self-contained treatment-specific Harbor environment.
# The build context contains only the selected treatment distribution.

FROM --platform=linux/amd64 python:3.13-slim@sha256:cc9dffa47c8294ba9bb795a8dfaeb7b76f2b30acade2c52a461a2999d127eb00

# Headless office runtime, identical across package and skill treatments.
RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      libreoffice-impress fonts-dejavu-core fonts-liberation poppler-utils \
    && soffice --version \
    && rm -rf /var/lib/apt/lists/*

ARG DEBIAN_FRONTEND=noninteractive

COPY wheelhouse/shared/ /tmp/pptx-shared-wheels/
RUN python -m pip install --no-cache-dir --no-index \
      --find-links=/tmp/pptx-shared-wheels \
      lxml==6.1.1 \
      Pillow==12.3.0 \
      typing_extensions==4.16.0 \
      XlsxWriter==3.2.9 \
    && rm -rf /tmp/pptx-shared-wheels

COPY images/common/pptx_treatment_fingerprint.py /usr/local/bin/pptx-ablation-fingerprint
COPY images/common/verify_treatment.sh /usr/local/bin/pptx-ablation-verify-treatment
COPY images/common/treatment_entrypoint.sh /usr/local/bin/pptx-ablation-entrypoint

RUN useradd --create-home --shell /bin/bash agent \
    && mkdir -p /app /opt/pptx-ablation \
    && chown agent:agent /app \
    && chmod 0755 \
      /usr/local/bin/pptx-ablation-fingerprint \
      /usr/local/bin/pptx-ablation-verify-treatment \
      /usr/local/bin/pptx-ablation-entrypoint

ENV PYTHONDONTWRITEBYTECODE=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_ROOT_USER_ACTION=ignore

WORKDIR /app
ENTRYPOINT ["/usr/local/bin/pptx-ablation-entrypoint"]

# Selected treatment layer
COPY wheelhouse/paper/ /tmp/pptx-treatment-wheels/
RUN python -m pip install --no-cache-dir --no-index \
      --find-links=/tmp/pptx-treatment-wheels \
      paper-pptx==0.2.0 \
    && rm -rf /tmp/pptx-treatment-wheels

ENV PPTX_ABLATION_CONDITION=paper
RUN python3 /usr/local/bin/pptx-ablation-fingerprint \
      --condition paper \
      > /opt/pptx-ablation/treatment-fingerprint.json

# Task fixture layer
WORKDIR /app
RUN mkdir -p /app/eval_fixtures/pptx /app/evals
COPY filesystem/eval_fixture_04_structure.pptx /app/eval_fixtures/pptx/
RUN chown -R agent:agent /app
