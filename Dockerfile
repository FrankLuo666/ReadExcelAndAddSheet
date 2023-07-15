# 使用一个基础的Java镜像
FROM adoptopenjdk/openjdk11:latest

# 设置工作目录
WORKDIR /app

# 复制项目文件到镜像中
COPY . /app

# 构建项目
RUN ./gradlew build

# 暴露Spring Boot应用的端口
EXPOSE 8080

# 启动Spring Boot应用
CMD ["java", "-jar", "build/libs/com/tool/ReadExcelAndAddSheet.jar"]