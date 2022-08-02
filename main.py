import pandas as pd
import numpy as np
import glob
import jpype
import jpype.imports
import xlsxwriter


class FormulaExecution:
    def __init__(self, df, formula_str):
        self.df = df
        self.formula_string = formula_str
        self.row = len(self.df)

    def formula_execution(self, node):
        if node.isLeafNode:
            rowStart = node.rowStart
            colStart = node.colStart
            rowEnd = node.rowEnd
            colEnd = node.colEnd
            startRelative = node.startRelative
            endRelative = node.endRelative
            if startRelative:
                ptype = 'R'
            else:
                ptype = 'F'
            if endRelative:
                ptype += 'R'
            else:
                ptype += 'F'
            return (((colStart, rowStart), (colEnd + 1, rowEnd + 1)), ptype)
        else:
            print(node.value)
            subtrees = []
            for child in node.children:
                subtree = self.formula_execution(child)
                subtrees.append(subtree)
            return self.compute_formula(subtrees, node.value)

    def compute_formula(self, nodes, operator):
        binary_ops = ['+', '-', '*', '/']
        if operator in binary_ops:
            return self.compute_binary_op(nodes, operator)
        if operator == 'SUM':
            return self.compute_sum(nodes)
        elif operator == 'AVERAGE':
            return self.compute_avg(nodes)
        elif operator == 'COUNT':
            return self.compute_count(nodes)
        elif operator == 'MAX':
            return self.compute_max(nodes)
        elif operator == 'MIN':
            return self.compute_min(nodes)
        else:
            print('Not supported')

    def compute_sum(self, nodes):
        values = []
        for node in nodes:
            if isinstance(node, tuple):
                ptype = node[1]
                head_col = node[0][0][0]
                tail_col = node[0][1][0]
                head_row = node[0][0][1]
                tail_row = node[0][1][1]
                h = tail_row - head_row
                df = self.df.iloc[head_row:, head_col:tail_col]
                if ptype == 'RR':
                    value = df.rolling(h).sum().dropna().sum(axis=1)
                elif ptype == 'FF':
                    value = np.array(df.iloc[head_row:tail_row]).sum()
                    value = np.full(self.row, value)
                elif ptype == 'FR':
                    value = df.expanding(h).sum().dropna().sum(axis=1)
                elif ptype == 'RF':
                    value = df.iloc[::-1].expanding(h - self.row + 1).sum().dropna().iloc[::-1].sum(axis=1)
                if ptype != 'FF' and self.row > len(value):
                    value = np.append(value, np.full(self.row - len(value), np.nan))
                values.append(value)
            else:
                values.append(node)
        return np.sum(values, axis=0)

    def compute_binary_op(self, nodes, op):
        df = self.df
        values = []
        for node in nodes:
            if isinstance(node, tuple):
                ptype = node[1]
                head_col = node[0][0][0]
                head_row = node[0][0][1]

                if ptype == 'RR':
                    value = df.iloc[head_row:, head_col]
                    if self.row > len(value):
                        value = np.append(value, np.full(self.row - len(value), np.nan))
                elif ptype == 'FF':
                    value = df.iloc[head_row, head_col]
                    value = np.full(self.row, value)
                values.append(value)
            else:
                values.append(node)
        if op == '+':
            return values[0] + values[1]
        elif op == '-':
            return values[0] - values[1]
        elif op == '*':
            return values[0] * values[1]
        elif op == '/':
            return values[0] / values[1]

    def compute_avg(self, nodes):
        values = []
        for node in nodes:
            if isinstance(node, tuple):
                ptype = node[1]
                head_col = node[0][0][0]
                tail_col = node[0][1][0]
                head_row = node[0][0][1]
                tail_row = node[0][1][1]
                h = tail_row - head_row
                df = self.df.iloc[head_row:, head_col:tail_col]
                if ptype == 'RR':
                    value = df.rolling(h).mean().dropna().mean(axis=1)
                elif ptype == 'FF':
                    value = np.array(df.iloc[head_row:tail_row]).mean()
                    value = np.full(self.row, value)
                elif ptype == 'FR':
                    value = df.expanding(h).mean().dropna().mean(axis=1)
                elif ptype == 'RF':
                    value = df.iloc[::-1].expanding(h - self.row + 1).mean().dropna().iloc[::-1].mean(axis=1)

                if ptype != 'FF'and self.row > len(value):
                    value = np.append(value, np.full(self.row - len(value), np.nan))
                values.append(value)
            else:
                values.append(node)
        return np.mean(values, axis=0)

    def compute_count(self, nodes):
        values = []
        for node in nodes:
            if isinstance(node, tuple):
                ptype = node[1]
                head_col = node[0][0][0]
                tail_col = node[0][1][0]
                head_row = node[0][0][1]
                tail_row = node[0][1][1]
                h = tail_row - head_row
                df = self.df.iloc[head_row:, head_col:tail_col]
                if ptype == 'RR':
                    value = df.rolling(h).count().dropna().sum(axis=1)
                elif ptype == 'FF':
                    value = df.iloc[head_row:tail_row].count().dropna().sum()
                    value = np.full(self.row, value)
                elif ptype == 'FR':
                    value = df.expanding(h).count().dropna().sum(axis=1)
                elif ptype == 'RF':
                    value = df.iloc[::-1].expanding(h - self.row + 1).count().dropna().iloc[::-1].sum(axis=1)
                if ptype != 'FF' and self.row > len(value):
                    value = np.append(value, np.full(self.row - len(value), np.nan))
                values.append(value)
            else:
                values.append(node)
        return np.sum(values, axis=0)

    def compute_max(self, nodes):
        values = []
        for node in nodes:
            if isinstance(node, tuple):
                ptype = node[1]
                head_col = node[0][0][0]
                tail_col = node[0][1][0]
                head_row = node[0][0][1]
                tail_row = node[0][1][1]
                h = tail_row - head_row
                df = self.df.iloc[head_row:, head_col:tail_col]
                if ptype == 'RR':
                    value = df.rolling(h).max().dropna().max(axis=1)
                elif ptype == 'FF':
                    value = np.array(df.iloc[head_row:tail_row]).max()
                    value = np.full(self.row, value)
                elif ptype == 'FR':
                    value = df.expanding(h).max().dropna().max(axis=1)
                elif ptype == 'RF':
                    value = df.iloc[::-1].expanding(h - self.row + 1).max().dropna().iloc[::-1].max(axis=1)
                if ptype != 'FF' and self.row > len(value):
                    value = np.append(value, np.full(self.row - len(value), np.nan))
                values.append(value)
            else:
                values.append(node)
        return np.max(values, axis=0)

    def compute_min(self, nodes):
        values = []
        for node in nodes:
            if isinstance(node, tuple):
                ptype = node[1]
                head_col = node[0][0][0]
                tail_col = node[0][1][0]
                head_row = node[0][0][1]
                tail_row = node[0][1][1]
                h = tail_row - head_row
                df = self.df.iloc[head_row:, head_col:tail_col]
                if ptype == 'RR':
                    value = df.rolling(h).min().dropna().min(axis=1)
                elif ptype == 'FF':
                    value = np.array(df.iloc[head_row:tail_row]).min()
                    value = np.full(self.row, value)
                elif ptype == 'FR':
                    value = df.expanding(h).min().dropna().min(axis=1)
                elif ptype == 'RF':
                    value = df.iloc[::-1].expanding(h - self.row + 1).min().dropna().iloc[::-1].min(axis=1)
                if ptype != 'FF' and self.row > len(value):
                    value = np.append(value, np.full(self.row - len(value), np.nan))
                values.append(value)
            else:
                values.append(node)
        return np.min(values, axis=0)

    def get_result(self):
        workbook = xlsxwriter.Workbook('workbook.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', formula_string)
        workbook.close()

        jars = glob.glob('sheetanalyzer_jar/*.jar')
        jpype.startJVM(classpath=':'.join(jars))

        # Import of Java classes must happen *after* jpype.startJVM() is called
        from org.dataspread.sheetanalyzer import SheetAnalyzer
        import org.dataspread.sheetanalyzer.parser.Node

        sheet = SheetAnalyzer.createSheetAnalyzer('workbook.xlsx')
        root = sheet.getFormulaTree()
        return self.formula_execution(root)


if __name__ == '__main__':
    m = 100
    n = 5
    df = pd.DataFrame(np.ones((m, n)))

    formula_string = "=SUM(A1:B2) / COUNT(A$2:A2)"
    fe = FormulaExecution(df, formula_string)
    print(fe.get_result())
