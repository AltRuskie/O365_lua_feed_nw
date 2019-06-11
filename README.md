# rsa_nw_lua_feed_o365_whitelist
whitelisting Office365 traffic using Lua and Feeds

https://community.rsa.com/community/products/netwitness/blog/2018/08/14/lua-whitelisting-office365-traffic

Data and sample script come from here
https://support.office.com/en-gb/article/managing-office-365-endpoints-99cab9d4-ef59-4207-9f2b-3728eb46bf9a?ui=en-US&rs=en-GB&ad=GB#ID0EACAAA=4._Web_service

This replaces the published xml feed that will EOL in October 2018

PS1 script publishes 3 outputs

o365ipv4Out.txt
o365ipv6Out.txt
o365UrlOut.txt

first two are used as feeds in NetWitness with the included xml files

url output is used to update the lua parser which does the url matching and wildcard checking.

the script keeps a tracker in the temp dir of what the latest version is and if there are any changes (also generates a clientID for you).
lua_O365.lua reads from alias.host and writes to filter metakeys (as well as the feeds)

filter='whitelist'
filter='office365'
