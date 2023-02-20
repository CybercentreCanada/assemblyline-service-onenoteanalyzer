ARG branch=latest
FROM cccs/assemblyline-v4-service-base:$branch

ENV SERVICE_PATH onenoteanalyzer.onenoteanalyzer.OneNoteAnalyzer

USER root

RUN apt-get update && apt-get install -y wget unzip

# Install Wine to run OneNoteAnalyzer (C# app using Aspose)
RUN dpkg --add-architecture i386 && mkdir -pm755 /etc/apt/keyrings && \
    wget -O /etc/apt/keyrings/winehq-archive.key https://dl.winehq.org/wine-builds/winehq.key
RUN wget -NP /etc/apt/sources.list.d/ https://dl.winehq.org/wine-builds/debian/dists/buster/winehq-buster.sources && \
    apt update && apt install -y --install-recommends winehq-stable
RUN wget https://github.com/knight0x07/OneNoteAnalyzer/releases/download/OneNoteAnalyzer/OneNoteAnalyzer.zip && \
    unzip OneNoteAnalyzer.zip -d /opt/al_service/OneNoteAnalyzer && rm -f OneNoteAnalyzer.zip
RUN wget -O /opt/al_service/dotNetFx40_Full_x86_x64.exe 'http://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe'...

# Switch to assemblyline user
USER assemblyline

# Copy service code
WORKDIR /opt/al_service
COPY . .

# Install dotnet under the AL user in Wine
RUN wine dotNetFx40_Full_x86_x64.exe /q

# Patch version in manifest
ARG version=4.0.0.dev1
USER root
RUN sed -i -e "s/\$SERVICE_TAG/$version/g" service_manifest.yml

# Switch to assemblyline user
USER assemblyline