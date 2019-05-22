#! /usr/bin/python
from table_fsm import Delegate, StateMachine

class ParsingContext:
  def __init__(self):
    self.buf = ""
    self.tmp = ""
    self.ch = None
    self.line = 1
    self.col = 1
    self.cells = []
    self.lines = []
    self.rows = []

class ParsingDelegate(Delegate):
  def error(self, ctx):
    print("Invalid table format at col %d in line %d" % (ctx.col, ctx.line))
    exit(-1)
  def append(self, ctx):
    ctx.buf += ctx.ch
  def cell(self, ctx):
    ctx.cells.append(ctx.buf.strip())
    ctx.buf = ''
  def line(self, ctx):
    ctx.lines.append(ctx.cells)
    ctx.cells = []
  def row(self, ctx):
    cells = []
    for i in range(len(ctx.lines[0])):
      cells.append([])
    for row in range(len(ctx.lines)):
      for col in range(len(ctx.lines[row])):
        if len(ctx.lines[row][col]) > 0:
            cells[col].append(ctx.lines[row][col])
    row = []
    for cell in cells:
      if len(cell) > 0:
        content = '\n'.join(cell)
        if content.startswith('----') or content.startswith('===='):
          content = '\n' + content
        if content.endswith('----') or content.endswith('===='):
          content += '\n'
        row.append(content)
      else:
        row.append(None)
    ctx.rows.append(row)
    ctx.lines = []

def load(src: str):
  ctx = ParsingContext()
  fsm = StateMachine(ParsingDelegate())
  with open(src, 'r') as input:
    content = input.read(-1)
    for ch in content:
      ctx.ch = ch
      if ch == '\n':
        ctx.line += 1
        ctx.col = 1
        fsm.lf(ctx)
      elif ch == '+':
        fsm.plus(ctx)
        ctx.col += 1
      elif ch == '-':
        fsm.minus(ctx)
        ctx.col += 1
      elif ch == '|':
        fsm.pipe(ctx)
        ctx.col += 1
      else:
        fsm.others(ctx)
        ctx.col += 1
  return ctx.rows

def save(model, dst: str):
  import xlsxwriter
  wb = xlsxwriter.Workbook(dst)
  ws = wb.add_worksheet()
  for row in range(len(model)):
    for col in range(len(model[row])):
      if model[row][col]:
        ws.write_string(row, col, model[row][col])
      else:
        ws.write_blank(row, col, None)
  wb.close()

def main(src: str, dst: str):
  model = load(src)
  save(model, dst)

if __name__ == '__main__':
  import argparse
  import sys
  parser = argparse.ArgumentParser()
  parser.add_argument("src", help="The table file")
  parser.add_argument("dst", help="The execel file")
  args = parser.parse_args()
  main(args.src, args.dst)
