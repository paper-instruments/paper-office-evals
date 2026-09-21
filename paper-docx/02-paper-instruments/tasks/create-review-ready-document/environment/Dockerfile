# Self-contained treatment-specific Harbor environment.
# The build context contains only the selected treatment distribution.

FROM python:3.13-slim@sha256:eb43ff125d8d58d7449dcba7d336c23bcac412f526d861db493b9994d8010280

# Headless office runtime, identical across package and skill treatments.
RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      libreoffice-writer fonts-dejavu-core fonts-liberation poppler-utils \
    && soffice --version \
    && rm -rf /var/lib/apt/lists/*

COPY wheelhouse /wheelhouse
RUN python -m pip install --no-cache-dir --no-index \
      --find-links=/wheelhouse/shared \
      lxml==6.1.1 typing_extensions==4.16.0

COPY images/common/docx_treatment_fingerprint.py /usr/local/bin/docx-ablation-fingerprint
COPY images/common/treatment_entrypoint.sh /usr/local/bin/docx-ablation-entrypoint
COPY images/common/verify_treatment.sh /usr/local/bin/docx-ablation-verify-treatment

RUN useradd --create-home --shell /bin/bash agent \
    && mkdir -p /app /opt/docx-ablation \
    && chown agent:agent /app \
    && chmod 0755 \
      /usr/local/bin/docx-ablation-fingerprint \
      /usr/local/bin/docx-ablation-entrypoint \
      /usr/local/bin/docx-ablation-verify-treatment

ENV PYTHONDONTWRITEBYTECODE=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_ROOT_USER_ACTION=ignore

WORKDIR /app
ENTRYPOINT ["/usr/local/bin/docx-ablation-entrypoint"]

# Selected treatment layer
RUN printf '%s  %s\n' 7fbed09dea52a1cf63b9f0afd39d75a6e843813a989feab247caffb1b0afa894 /wheelhouse/paper/paper_docx-0.2.0-py3-none-any.whl | sha256sum -c -
RUN python -m pip install --no-cache-dir --no-index --no-deps \
      /wheelhouse/paper/paper_docx-0.2.0-py3-none-any.whl

ENV DOCX_ABLATION_CONDITION=paper
RUN printf '%s\n' paper > /opt/docx-ablation/condition \
    && python3 /usr/local/bin/docx-ablation-fingerprint \
      --expected paper \
      --output /opt/docx-ablation/treatment-fingerprint.json

# Task fixture layer
WORKDIR /app
RUN mkdir -p /app/eval_fixtures/docx /app/evals
RUN chown -R agent:agent /app
