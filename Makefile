#
#
all:

deploy:
	rsync \
		--exclude=Makefile \
		--exclude=MK.sh \
		--exclude=.gitignore \
		--exclude-from=.gitignore \
	--delete -av . nuws@msa:nuws/msa-ps1/
