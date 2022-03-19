rm -f *.aiff
rm -f *.mp3
say -o today -f raw.txt
lame today.aiff today.mp3
