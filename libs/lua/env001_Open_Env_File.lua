local f = io.open(".env", "r")
local contents = f:read("*all")
f:close()

print("Contents of .env file:")
print(contents)
