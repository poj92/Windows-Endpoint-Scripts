# Open Azure CLI and switch to Bash
# Now make Azure CLI to popup login details to allow login to the tenant
export AZCOPY_AUTO_LOGIN_TYPE=DEVICE

# NOTE
    # If moving from one tenant to the other, it is best to login to the source tenant when prompted and generate a SAS from the 
    # destination tenant to append to the url of the account/container/blob
#

# Determine IP of CLoud Shell to use to generate SAS
curl -s checkip.dyndns.org | sed -e 's/.*Current IP Address: //' -e 's/<.*$//'

# 
# You can check that you have the access that is required by running the following to List files 
azcopy list https://dacorumfilestore.blob.core.windows.net/

# Copy account
azcopy copy 'https://gasdocuments.blob.core.windows.net/' 'https://cardogasdocuments.blob.core.windows.net/?sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2024-01-29T19:29:57Z&st=2024-01-29T11:29:57Z&sip=20.31.242.253&spr=https&sig=NTRKPxrwR904PFtvg5sOjjG3ZX%2BnR3jDMkixosTnbw8%3D' --recursive



az storage file download --path '/home/nexus/.azcopy/2eb5f718-1663-234b-50f9-083da7bc47ea.log'

az storage file download --path '/home/nexus/.azcopy/2eb5f718-1663-234b-50f9-083da7bc47ea.log'