/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2025 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2025 UNVELL Inc. All rights reserved.
 * 
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace unvell.ReoGrid.IO
{
	internal static class CSVFormat
	{
		public const int DEFAULT_READ_BUFFER_LINES = 512;

		public static void Read(Stream stream, Worksheet sheet, RangePosition targetRange,
			Encoding? encoding = null, int bufferLines = DEFAULT_READ_BUFFER_LINES, bool autoSpread = true)
		{
			targetRange = sheet.FixRange(targetRange);

			List<object>[] bufferLineList = new List<object>[bufferLines];

			for (int i = 0; i < bufferLineList.Length; i++)
			{
				bufferLineList[i] = new List<object>(256);
			}

#if DEBUG
			var sw = System.Diagnostics.Stopwatch.StartNew();
#endif

			int row = targetRange.Row;
			int totalReadLines = 0;

			using (var sr = new StreamReader(stream, encoding))
			{
				var records = new CSVRecordReader(sr);

				sheet.SuspendDataChangedEvents();
				int maxCols = 0;

				try
				{
					bool finished = false;
					while (!finished)
					{
						int readLines = 0;

						for (; readLines < bufferLineList.Length; readLines++)
						{
							if (!records.ReadRecord(bufferLineList[readLines], targetRange.Cols))
							{
								finished = true;
								break;
							}

							totalReadLines++;
							if (!autoSpread && totalReadLines > targetRange.Rows)
							{
								finished = true;
								break;
							}
						}

						if (autoSpread && row + readLines > sheet.RowCount)
						{
							int appendRows = bufferLines - (sheet.RowCount % bufferLines);
							if (appendRows <= 0) appendRows = bufferLines;
							sheet.AppendRows(appendRows);
						}

						for (int i = 0; i < readLines; i++)
						{
							if (maxCols < bufferLineList[i].Count) maxCols = bufferLineList[i].Count;

							if (autoSpread && maxCols >= sheet.ColumnCount)
							{
								sheet.SetCols(maxCols + 1);
							}
						}

						sheet.SetRangeData(row, targetRange.Col, readLines, maxCols, bufferLineList);
						row += readLines;
					}
				}
				finally
				{
					sheet.ResumeDataChangedEvents();
				}

				sheet.RaiseRangeDataChangedEvent(new RangePosition(
					targetRange.Row, targetRange.Col, maxCols, totalReadLines));
			}

#if DEBUG
			sw.Stop();
			System.Diagnostics.Debug.WriteLine("load csv file: " + sw.ElapsedMilliseconds + " ms, rows: " + row);
#endif
		}
	}

	/// <summary>
	/// Reads CSV records as RFC 4180 describes them: a quoted field may hold separators, line
	/// breaks and doubled quotes, and every other character belongs to the field it is in.
	/// </summary>
	internal sealed class CSVRecordReader
	{
		private const int NothingPeeked = -2;

		private readonly TextReader reader;
		private readonly StringBuilder field = new StringBuilder();
		private int peeked = NothingPeeked;

		public CSVRecordReader(TextReader reader)
		{
			this.reader = reader;
		}

		/// <summary>
		/// Read the next record, keeping at most the specified number of fields.
		/// </summary>
		/// <returns>False when the input has no more records.</returns>
		public bool ReadRecord(List<object> fields, int maxFields)
		{
			fields.Clear();

			int c = Read();
			if (c == -1) return false;

			bool inQuotes = false;
			bool fieldStart = true;

			while (true)
			{
				if (inQuotes)
				{
					// An unterminated quote runs to the end of the input.
					if (c == -1) break;

					if (c == '"')
					{
						if (Peek() == '"')
						{
							Read();
							field.Append('"');
						}
						else
						{
							inQuotes = false;
						}
					}
					else
					{
						field.Append((char)c);
					}
				}
				else if (c == '"' && fieldStart)
				{
					inQuotes = true;
				}
				else if (c == ',')
				{
					AddField(fields, maxFields);
					fieldStart = true;
					c = Read();
					continue;
				}
				else if (c == '\r' || c == '\n' || c == -1)
				{
					if (c == '\r' && Peek() == '\n') Read();
					break;
				}
				else
				{
					field.Append((char)c);
				}

				fieldStart = false;
				c = Read();
			}

			AddField(fields, maxFields);
			return true;
		}

		private void AddField(List<object> fields, int maxFields)
		{
			if (fields.Count < maxFields) fields.Add(field.ToString());
			field.Clear();
		}

		private int Read()
		{
			if (peeked == NothingPeeked) return reader.Read();

			int c = peeked;
			peeked = NothingPeeked;
			return c;
		}

		private int Peek()
		{
			if (peeked == NothingPeeked) peeked = reader.Read();
			return peeked;
		}
	}

	#region CSV File Provider
	internal class CSVFileFormatProvider : IFileFormatProvider
	{
		public bool IsValidFormat(string file)
		{
			return System.IO.Path.GetExtension(file).Equals(".csv", StringComparison.CurrentCultureIgnoreCase);
		}

		public bool IsValidFormat(Stream s)
		{
			throw new NotSupportedException();
		}

		public void Load(IWorkbook workbook, Stream stream, Encoding? encoding, object? arg)
		{
			bool autoSpread = true;
			int bufferLines = CSVFormat.DEFAULT_READ_BUFFER_LINES;
			RangePosition targetRange = RangePosition.EntireRange;

			CSVFormatArgument? csvArg = arg as CSVFormatArgument;

			if (csvArg != null)
			{
				autoSpread = csvArg.AutoSpread;
				bufferLines = csvArg.BufferLines;
				targetRange = csvArg.TargetRange;
			}
		
			Worksheet? sheet = null;

			if (workbook.Worksheets.Count == 0)
			{
				sheet = workbook.CreateWorksheet("Sheet1");
				workbook.Worksheets.Add(sheet);
			}
			else
			{
				while (workbook.Worksheets.Count > 1)
				{
					workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);
				}

				sheet = workbook.Worksheets[0];
				sheet.Reset();
			}

			CSVFormat.Read(stream, sheet, targetRange, encoding, bufferLines, autoSpread);
		}

		public void Save(IWorkbook workbook, Stream stream, Encoding? encoding, object? arg)
		{
			throw new NotSupportedException("Saving entire workbook as CSV is not supported, use Worksheet.ExportAsCSV instead.");
			//int fromRow = 0, fromCol = 0, toRow = 0, toCol = 0;

			//if (args != null)
			//{
			//	object arg;
			//	if (args.TryGetValue("fromRow", out arg)) fromRow = (int)arg;
			//	if (args.TryGetValue("fromCol", out arg)) fromCol = (int)arg;
			//	if (args.TryGetValue("toRow", out arg)) toRow = (int)arg;
			//	if (args.TryGetValue("toCol", out arg)) toCol = (int)arg;
			//}
		}
	}

	/// <summary>
	/// Arguments for loading and saving CSV format.
	/// </summary>
	public class CSVFormatArgument
	{
		/// <summary>
		/// Determines whether or not allow to expand worksheet to load more data from CSV file. (Default is True)
		/// </summary>
		public bool AutoSpread { get; set; }

		/// <summary>
		/// Determines how many rows read from CSV file one time. (Default is CSVFormat.DEFAULT_READ_BUFFER_LINES = 512)
		/// </summary>
		public int BufferLines { get; set; }

		/// <summary>
		/// Determines the default worksheet name if CSV file loaded into a new workbook. (Default is "Sheet1")
		/// </summary>
		public string SheetName { get; set; }

		/// <summary>
		/// Determines where to start import CSV data on worksheet.
		/// </summary>
		public RangePosition TargetRange { get; set; }

		/// <summary>
		/// Create the argument object instance.
		/// </summary>
		public CSVFormatArgument()
		{
			this.AutoSpread = true;
			this.BufferLines = CSVFormat.DEFAULT_READ_BUFFER_LINES;
			this.SheetName = "Sheet1";
			this.TargetRange = RangePosition.EntireRange;
		}
	}
	#endregion // CSV File Provider

}
