####  从excel中读取数据并增加sheet，并在其中填入读取的数据:

 最开始，删除出了sheet1和sheet2之外的所有sheet。
 从第1个sheet的第23行开始，取出这一行几个字段第值。
 每取得1行，就新建一个sheet，以sheet2为模板，将从sheet1取得的值，插入到新sheet的指定cell中。
 最后删除模板sheet2。

####  sheet1的几个字段及其值的例子：
 num	functionId	modifier	functionName(logical)	functionName(physical)
 1	001	public	メソード１	function1
 2	002	private	メソード２	function2
 3	003	private	メソード３	function3


<br/><br/>

#### 使用Dockerfile建立docker镜像
```shell
docker build -f ./Dockerfile -t ubuntu_jdk17_maven:1 .
```
#### 查看生成的镜像
```shell
docker images
```

#### 启动镜像
```shell
docker run -it --name=ReadExcelAndAddSheet ubuntu_jdk17_maven:1 bash
```

#### 在容器中启动项目
```shell
sh run.sh
```

<br/><br/>

### 如果需要主机和容器共享文件夹
```shell
# 创建镜像：docker build -t <镜像名称> <Dockerfile路径>
docker build -t ubuntu_jdk17_maven:1 .

# 启动并进入容器： docker run -it -v /绝对路径/shared:/app/data <其他参数> <镜像名称> /bin/bash
docker run -it -v /Users/luo/dockerShared:/app/data ubuntu_jdk17_maven:1 /bin/bash
```
- docker build -t <镜像名称> <Dockerfile路径>：
这个命令用于构建 Docker 镜像。通过指定 -t 参数和镜像名称，你可以为镜像指定一个易于识别的名称。<Dockerfile路径> 是 Dockerfile 文件的路径，它包含了构建镜像所需的指令和配置。

- docker run -it -v /绝对路径/shared:/app/data <其他参数> <镜像名称> /bin/bash
这个命令用于运行Docker容器并进入终端。-v 参数用于创建容器内的挂载点，将宿主机的目录 /绝对路径/shared 挂载到容器内的 /app/data 目录上，实现了数据的共享。<其他参数> 是可选的，可以包含其他运行容器所需的配置参数。<镜像名称> 是你要运行的镜像的名称。
  - -it 表示以交互模式运行容器，并分配一个终端。
  - -v /绝对路径/shared:/app/data 指定主机文件系统上的绝对路径与容器内的 /app/data 目录之间建立挂载点，实现数据共享。
  - <其他参数> 可以是您希望传递给容器的其他选项或参数。
  - <镜像名称> 是您要运行的 Docker 镜像的名称。
  - /bin/bash 是容器中要执行的命令，这将在容器启动后进入交互式终端。


这两个命令结合使用，首先构建镜像，然后运行该镜像的容器，并通过挂载目录实现宿主机和容器之间的数据共享。


### 提交镜像
1. 使用docker login命令登录到远程仓库。例如，如果您使用的是 Docker Hub，可以运行以下命令登录：
```shell
docker login
```
您将被要求输入您的用户名和密码。

2. 标记您的本地镜像，以指定远程仓库的名称和标签。使用docker tag命令，将镜像重新命名为远程仓库的名称。例如，如果您的本地镜像名称是 my-image，远程仓库是 Docker Hub，您可以运行以下命令：
```shell
# docker tag <本地镜像>:<标签> <仓库用户名>/<仓库名称>:<标签>
docker tag ubuntu_jdk17_maven:1 frankluo666/ubuntu_jdk17_maven:1
```
请将 <仓库用户名> 替换为您的 Docker Hub 用户名，<仓库名称> 替换为您的仓库名称，<标签> 替换为您想要的标签。

3. 使用docker push命令将标记的镜像推送到远程仓库。例如，如果您的镜像标签是 latest，您可以运行以下命令：
```shell
# docker push <仓库名称>:<标签>
docker push ubuntu_jdk17_maven:1
```
这将上传您的镜像到远程仓库。

完成上述步骤后，您的镜像将被推送到远程仓库，并可在其他地方访问和使用。请确保已正确替换命令中的占位符，并根据您使用的远程仓库服务进行相应的登录和配置。


