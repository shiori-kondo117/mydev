if [ $# -ne 2 ];then
	echo "Usage:$(basename $0) [drive-id] [src-dir]"
	exit 1;
fi
DATA_ID=`tail -1 /cygdrive/j/mydata/db/csv/vr_titles_hdrv.csv|awk -F, '{ print $1; }'`;
find /cygdrive/$1/$2 -name "VR_MANGR.IFO" -exec ~/bin/get_vr_title_lists.sh {} \;|awk -v DATA_ID=${DATA_ID} '{ ++DATA_ID;printf("%s,%s\n",DATA_ID,$0); }' >> /cygdrive/m/mydata/db/csv/vr_titles_hdrv.csv
