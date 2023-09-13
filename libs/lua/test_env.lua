local f = io.open(".env", "r")
local contents = f:read("*all")
f:close()
print("Contents of .env file:")
print(contents)
-- 打开 .env 文件
local f = io.open(".env", "r")

if not f then
	return nil, ".env 文件不存在或无法打开"
end

-- 初始化变量
local variableName = "IMAGE_URL"
local variableValue = nil

-- 遍历文件的每一行
for line in f:lines() do
	local match_value = string.match(line, "=(.*)$")
	print("match_value:", match_value)
end

-- 关闭文件
f:close()
