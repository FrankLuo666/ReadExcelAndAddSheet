# 使用基础 Linux 镜像
FROM ubuntu:latest

# 安装 Java JDK 17
RUN apt-get update && \
    apt-get install -y openjdk-17-jdk

# 安装 Maven
RUN apt-get install -y maven

# 安装 Git
RUN apt-get install -y git

# 安装 Vim
RUN apt-get install -y vim

# 设置环境变量
ENV JAVA_HOME /usr/lib/jvm/java-17-openjdk-amd64
ENV MAVEN_HOME /usr/share/maven
ENV PATH $PATH:$JAVA_HOME/bin:$MAVEN_HOME/bin

# 设置工作目录
WORKDIR /app

# 下载代码
RUN git clone https://github.com/FrankLuo666/ReadExcelAndAddSheet.git

# 设置工作目录到项目根目录
WORKDIR /app/ReadExcelAndAddSheet

# 构建项目
RUN mvn clean install

# 定义容器启动命令
CMD ["/bin/bash"]
