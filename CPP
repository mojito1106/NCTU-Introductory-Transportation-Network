import xlrd
import time

start = time.time()
N = 18
SOURCE = 0
GRAPH = []
DP = [[-1 for j in range(0, N)] for i in range(0, 1 << N)]
'''
for i in range(0, 1 << N):
    for j in range(0, N):
        DP[i][j] = -1
    '''
PATH = [[-1 for j in range(0, N)] for i in range(0, 1 << N)]


class TSP:
    @staticmethod
    def init(filename='距離矩陣.xlsx') -> None:
        book = xlrd.open_workbook(filename)
        sheet1 = book.sheets()[0]
        global LOCATION
        LOCATION = sheet1.row_values(0)[3:]
        for i in range(0, N):
            sub_graph = []
            for j in range(0, N):
                if j == i:
                    sub_graph.append(0)
                else:
                    sub_graph.append(sheet1.row_values(i + 1)[j + 3])
            GRAPH.append(sub_graph)

        for i in range(0, N):
            DP[1 << i][i] = GRAPH[SOURCE][i]

    @staticmethod
    def run(status: int, x: int) -> int:
        if DP[status][x] != -1:
            return DP[status][x]

        mask = 1 << x
        DP[status][x] = 1e9
        for i in range(0, N):
            if i != x and (status & (1 << i)):
                tmp = TSP.run(status - mask, i) + GRAPH[i][x]
                if tmp < DP[status][x]:
                    DP[status][x] = tmp
                    PATH[status][x] = i
        return DP[status][x]

    @staticmethod
    def print_result():
        # print PATH
        state = (1 << N) - 1
        x = SOURCE
        for i in range(0, N):
            print(LOCATION[x], end="<-")
            mask = 1 << x
            x = PATH[state][x]
            state = state - mask
        print(LOCATION[SOURCE])


if __name__ == '__main__':
    TSP.init()
    print('Solution:', TSP.run((1 << N) - 1, SOURCE))
    TSP.print_result()
    print('Takes: %f s' % (time.time() - start))
