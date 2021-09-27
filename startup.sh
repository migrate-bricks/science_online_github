#!/bin/bash

git_dir=./public
dev_dir=./dev

#create new folder for dev
if [ -d $dev_dir ]; then rm -Rf $dev_dir; fi
mkdir $dev_dir

if [ -d $git_dir ]; then rm -Rf $git_dir; fi
git clone https://github.com/pojiezhiyuanjun/freev2.git ./public

#search at most 30 days to find the latest v2ray file  
#an copy and rename to the destination folder

for i in {0..30}
do
	long_date=$(date +%Y%""m%d -d "$DATE - $i day")
	echo "long_date=$long_date"
	#filename=$(find ./*${long_date}*.yml | sed 's#.*/##')
	filepath=$(find $git_dir/*${long_date}*.yml)
	echo "filepath=$filepath"
	if [[ -f "$filepath" ]]
	then
		echo "This file exists on your filesystem."
		cp $filepath $dev_dir/free_node.yml
		break
	fi
	
	short_date=$(date +%m%d -d "$DATE - $i day")
	echo "short_date=$short_date"
	#filename=$(find ./*${short_date}*.yml | sed 's#.*/##')
	filepath=$(find $git_dir/*${short_date}*.yml)
	echo "filepath=$filepath"
	if [[ -f "$filepath" ]]
	then
		echo "This file exists on your filesystem."
		cp $filepath $dev_dir/free_node.yml
		break
	fi
	
done

if [ -d $git_dir ]; then rm -Rf $git_dir; fi

cd $dev_dir
git init
git config user.name "deployment bot"
git config user.email "deploy@github.com"
git add .
git commit -m "deploy node"
git config --global --unset http.proxy
git config --global --unset https.proxy
git push --force --quiet "https://${GH_TOKEN}@github.com/migrate-bricks/science_online_github.git" master:dev
#git push --force --quiet "https://oauth2:${GITEE_TOKEN_CLASH}@gitee.com/LinRaise/science_online_gitee.git" master:master
echo "complete!"
