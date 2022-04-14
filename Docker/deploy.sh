cd /Users/zhubo/code/zboDoMess/Docker/dockerFile &&
docker stop mess_d &&
docker rm mess_d
docker rmi zbodomess
cd /Users/zhubo/code/zboDoMess/Docker/dockerFile &&
docker build -t zbodomess .
docker run -d -p 5000:5000 -v /Users/zhubo/code/zboDoMess:/zboDoMess/ --name mess_d zbodomess