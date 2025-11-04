# 使用微软官方的.NET开发环境
FROM mcr.microsoft.com/devcontainers/dotnet:6.0

# 安装额外的VSTO开发工具
RUN apt-get update && export DEBIAN_FRONTEND=noninteractive \
    && apt-get -y install --no-install-recommends \
        git \
        zip \
        unzip \
        wget \
        curl \
        gnupg2 \
        software-properties-common \
        apt-transport-https \
        ca-certificates \
        build-essential \
        libssl-dev \
        libffi-dev \
        libxml2-dev \
        libxslt1-dev \
        zlib1g-dev \
        python3 \
        python3-pip \
        nodejs \
        npm \
    && rm -rf /var/lib/apt/lists/*

# 安装Mono（用于VSTO兼容性）
RUN apt-key adv --keyserver hkp://keyserver.ubuntu.com:80 --recv-keys 3FA7E0328081BFF6A14DA29AA6A19B38D3D831EF \
    && echo "deb https://download.mono-project.com/repo/ubuntu stable-focal main" | tee /etc/apt/sources.list.d/mono-official-stable.list \
    && apt-get update \
    && apt-get install -y mono-complete mono-xsp4 \
    && rm -rf /var/lib/apt/lists/*

# 安装NuGet
RUN wget -O /tmp/nuget.exe https://dist.nuget.org/win-x86-commandline/latest/nuget.exe \
    && mkdir -p /usr/local/bin \
    && mv /tmp/nuget.exe /usr/local/bin/nuget.exe

# 创建开发用户
ARG USERNAME=vscode
ARG USER_UID=1000
ARG USER_GID=$USER_UID
RUN if [ "$USER_GID" != "1000" ]; then groupmod --gid $USER_GID $USERNAME; fi \
    && if [ "$USER_UID" != "1000" ]; then usermod --uid $USER_UID --gid $USER_GID $USERNAME; fi

# 设置环境变量
ENV PATH="/usr/local/bin:${PATH}"
ENV DOTNET_ROOT=/usr/share/dotnet
ENV PATH=$PATH:/usr/local/bin/dotnet

# 清理
RUN apt-get autoremove -y && apt-get clean -y

WORKDIR /workspaces