using System;
using System.Collections;
using System.Collections.Generic;

namespace FastXL
{
	public readonly struct Row : IReadOnlyList<Cell>
	{
		readonly Cell[] cells;

		public Cell this[int index]
		{
			get => cells[index];
			internal set => cells[index] = value;
		}

		public int Count => cells.Length;

		internal Row(Cell[] cells)
		{
			this.cells = cells ?? throw new ArgumentNullException(nameof(cells));
		}

		public IEnumerator<Cell> GetEnumerator()
		{
			foreach (var cell in cells)
				yield return cell;
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return cells.GetEnumerator();
		}
	}
}
