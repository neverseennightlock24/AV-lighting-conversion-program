def parse_csv(filename):
    string = ""
    with open(filename, "r") as f:
        string += f.read()

    if len(string) == 0:
        return []

    ret1 = string.split("\n")
    retval = []
    for row in ret1:
        retval.append(row.split(","))

    f.close()
    return retval

fp = input("path: ")
filename = input("output: ")
output = open(filename, "w")
x = parse_csv(fp)
prev = False
for i in range(len(x)):
    if x[i][0] == "START_TARGETS":
        prev = True
        value = i
        break

value += 2
count = 0
while x[value][0] != "END_TARGETS":
    follow = float(x[value][23][1:])
    if count % 2 == 1:
        follow -= 0.01
        round(follow,3)
        newfollow = "F" + "%.2f"%follow
        x[value][23] = newfollow
    else:
        newfollow = follow
    print(follow)
    value += 1
    count += 1

rowed_grid = []
for row in x:
    rowed_grid.append(','.join(row))
parsed_grid = '\n'.join(rowed_grid)
output.write(parsed_grid)
output.close()
