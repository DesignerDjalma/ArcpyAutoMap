m1 = [
    [1, 2],
    [3, 4],
]
m2 = [
    [5, 4],
    [3, 2],
]


def somarMatrizes(matrix_A: list[list[int]], matrix_B: list[list[int]]):

    matrix_soma: list[list[int]] = []

    for linhas_A, linhas_B in zip(matrix_A, matrix_B):
        print(f"linhas_A: {linhas_A}")
        print(f"linhas_B: {linhas_B}")

        linhas_soma: list[int] = []

        for item_A, item_B in zip(linhas_A, linhas_B):
            print(f"item_A: {item_A}")
            print(f"item_B: {item_B}")
            linhas_soma.append(item_A + item_B)

        matrix_soma.append(linhas_soma)

    return matrix_soma


def somarM(m1, m2):
    mf = []  # matriz final
    for i, j in zip(m1, m2):
        linha = []
        for x, y in zip(i, j):
            linha.append(x + y)
        mf.append(linha)
    return mf


mf = somarM(m1, m2)
print(mf)
