const ExcelJS = require('exceljs');

// Define border styles
const DEFAULT_BORDER_STYLE = { style: 'solid', width: '1px' };
const BORDER_STYLES = {
	dashDot: { style: 'dashed', dash: 'dot' },
	dashDotDot: { style: 'dashed', dash: 'dotdot' },
	dashed: { style: 'dashed' },
	dotted: { style: 'dotted' },
	double: { style: 'double' },
	hair: { style: 'solid', width: '0.5px' },
	medium: { style: 'solid', width: '2px' },
	mediumDashDot: { style: 'dashed', width: '2px', dash: 'dot' },
	mediumDashDotDot: { style: 'dashed', width: '2px', dash: 'dotdot' },
	mediumDashed: { style: 'dashed', width: '2px' },
	slantDashDot: { style: 'dashed', dash: 'dot' },
	thick: { style: 'solid', width: '3px' },
	thin: { style: 'solid', width: '1px' }
};

const themeColors = [
	{ theme: 0, rgb: 'FF000000' }, // Black
	{ theme: 1, rgb: 'FFFFFFFF' }, // White
	{ theme: 2, rgb: 'FFFF0000' }, // Red
	{ theme: 3, rgb: 'FF00FF00' }, // Green
	{ theme: 4, rgb: 'FF0000FF' }, // Blue
	{ theme: 5, rgb: 'FFFF00FF' }, // Magenta
	{ theme: 6, rgb: 'FF00FFFF' }, // Cyan
	{ theme: 7, rgb: 'FFFFFF00' }, // Yellow
	{ theme: 8, rgb: 'FF800000' }, // Dark Red
	{ theme: 9, rgb: 'FF008000' }, // Dark Green
	{ theme: 10, rgb: 'FF000080' }, // Dark Blue
	{ theme: 11, rgb: 'FF800080' }, // Purple
	{ theme: 12, rgb: 'FF008080' }, // Teal
];

// Function to apply tint to an RGB color
function applyTint(rgb, tint) {
	const r = parseInt(rgb.substr(0, 2), 16);
	const g = parseInt(rgb.substr(2, 2), 16);
	const b = parseInt(rgb.substr(4, 2), 16);

	const newR = Math.round(r + (255 - r) * tint);
	const newG = Math.round(g + (255 - g) * tint);
	const newB = Math.round(b + (255 - b) * tint);
	return `${(newR.toString(16)).padStart(2, '0')}${(newG.toString(16)).padStart(2, '0')}${(newB.toString(16)).padStart(2, '0')}`;
}

// Function to convert theme color to RGB, considering tint
function getRgbFromTheme(themeIndex, tint = 0) {
	const themeColor = themeColors.find(color => color.theme === themeIndex);
	if (!themeColor) {
		throw new Error(`Theme color with index ${themeIndex} is not defined.`);
	}
	let rgb = themeColor.rgb.substr(2); // Remove the 'FF' part which is for alpha channel
	rgb = applyTint(rgb, tint);
	return `FF${rgb}`; // Adding the alpha channel back
}


// Convert ARGB color to hex
const argbToHex = (argb, theme, tint=0) => {
	if (theme) return `#${getRgbFromTheme(theme, tint).slice(2)}`; // can work uncorrect
	if (argb) return `#${argb.slice(2)}`;
	return null
};



// Get border styles from cell
const getBorderStyleFromCell = (cell) => {
	const borderStyles = {};
	['top', 'bottom', 'left', 'right'].forEach(side => {
		const border = cell.border?.[side];
		if (border && border.style) {
			const style = BORDER_STYLES[border.style] || DEFAULT_BORDER_STYLE;
			borderStyles[`border-${side}`] = `${style.width || DEFAULT_BORDER_STYLE.width} ${style.style}`;
			if (border.color) {
				borderStyles[`border-${side}-color`] = argbToHex(border.color.argb, null) || 'black';
			}
		}
	});
	return borderStyles;
};

// Get styles from cell
const getStylesFromCell = (cell) => {
	const styles = { 'border-collapse': 'collapse' };
	const borderStyles = getBorderStyleFromCell(cell);

	Object.assign(styles, borderStyles);

	if (cell.alignment) {
		if (cell.alignment.horizontal) styles['text-align'] = cell.alignment.horizontal;
		if (cell.alignment.vertical) styles['vertical-align'] = cell.alignment.vertical;
	}

	if (cell.fill?.fgColor) {
	 if (cell.fill.fgColor.argb || cell.fill.fgColor.theme)
		styles['background-color'] = argbToHex(cell.fill.fgColor.argb, cell.fill.fgColor.theme, cell.fill.fgColor.tint);
	
	}

	if (cell.style && cell.style.font) {
		if (cell.style.font.size) styles['font-size'] = `${cell.style.font.size}px`;
		const fontColor = argbToHex(cell.font.color?.argb, null);
		if (fontColor) styles['color'] = fontColor; // Only set color if it's not null
		if (cell.style.font.bold) styles['font-weight'] = 'bold';
		if (cell.style.font.italic) styles['font-style'] = 'italic';
		if (cell.style.font.underline) styles['text-decoration'] = 'underline';
		if (cell.style.font.name) styles['font-family'] = cell.style.font.name;
	}

	return styles;
};

// Process worksheet data
const worksheetToData = async (worksheet) => {
	const mergedCellMap = {};
	const excludedCells = new Set();


	// Process merged cells
	Object.keys(worksheet._merges).forEach(cellAddress => {
		try {
			const { model } = worksheet._merges[cellAddress];
			const { top, left, bottom, right } = model;

			const startCellAddress = worksheet.getCell(top, left).address;
			const endCellAddress = worksheet.getCell(bottom, right).address;
			const colspan = right - left + 1;
			const rowspan = bottom - top + 1;
			// Map start cell to its merged attributes and style
			mergedCellMap[startCellAddress] = {
				attrs: {
					colspan: colspan > 1 ? colspan : 1,
					rowspan: rowspan > 1 ? rowspan : 1,
				},
				style: {...getStylesFromCell(worksheet.getCell(bottom, right)), ...getStylesFromCell(worksheet.getCell(top, left)) }
			};

			// Add all cells in the merged range to the excludedCells set
			for (let row = top; row <= bottom; row++) {
				for (let col = left; col <= right; col++) {
					excludedCells.add(`${col}-${row}`);
				}
			}
		} catch (error) {
			console.error(`Error processing merge range:`, error);
		}
	});

	const data = [];
	const columns = [];

	worksheet.eachRow( (row, rowNumber) => {
		const dataRow = [];
		data.push(dataRow);
		row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
			const cellAddress = cell.address;
			if (!(
				excludedCells.has(`${colNumber}-${rowNumber}`) && 
				!mergedCellMap[cellAddress]
			) ){

				const cellData = {
					column: colNumber,
					row: rowNumber,
					value: cell.value || '',
					formattedValue: cell.value || '',
					attrs: {
						id: `${worksheet.name}!${cellAddress}`,
						...(mergedCellMap[cellAddress] ? mergedCellMap[cellAddress].attrs : {}),
					},
					style: mergedCellMap[cellAddress] ? mergedCellMap[cellAddress].style : getStylesFromCell(cell),
				};

				// Set cell styles
				cellData.style.height = `${worksheet.getRow(rowNumber).height || 19}px`;
				dataRow.push(cellData);
			} 
		});
	});

	const nonEmptyColumns = new Set();  // Набор для отслеживания непустых столбцов

	// Первый проход: определяем, какие столбцы непустые
	worksheet.eachRow((row, rowNumber) => {
		row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
			const cellValue = cell.value || '';
        const cellStyles = getStylesFromCell(cell);

        // Проверяем, есть ли значение или стили у ячейки, или она является частью объединённой ячейки
        if (cellValue !== '' || mergedCellMap[cell.address] || Object.keys(cellStyles).length > 0) {
            nonEmptyColumns.add(colNumber);
        }
			/*if (cellValue !== '' || mergedCellMap[cell.address]) {
				nonEmptyColumns.add(colNumber);
			}*/
		});
	});

	// Формируем массив columns, исключая пустые столбцы
	worksheet.columns.forEach((col, index) => {
		const colIndex = index + 1;

		// Добавляем столбец только если он не пустой
		if (nonEmptyColumns.has(colIndex)) {
			columns.push({
				index: colIndex,
				hidden: col.hidden,
				style: { width: `${((col.width || 10) * 10)}px` },
			});
		}
	});
	return { rows: data, cols: columns };
};

// Render table
const renderTable = (data) => {
	const htmlContent = [];

	htmlContent.push('<table style="border-collapse: collapse" border="0" cellspacing="1" cellpadding="0">');
	htmlContent.push('<colgroup>');
	data.cols.forEach(col => {
		if (!col.hidden) {
			htmlContent.push(`<col style="width: ${col.style.width}">`);
		}
	});
	htmlContent.push('</colgroup>');

	data.rows.forEach(row => {
		htmlContent.push('<tr>');
		row.forEach(cell => {
			const styleString = Object.entries(cell.style).map(([key, value]) => `${key}: ${value}`).join('; ');
			if (cell.attrs.colspan || cell.attrs.rowspan) {
				htmlContent.push(
					`<td ${cell.attrs.colspan ? `colspan="${cell.attrs.colspan}"` : '1'} ${cell.attrs.rowspan ? `rowspan="${cell.attrs.rowspan}"` : '1'} style="${styleString}">${cell.formattedValue || ''}</td>`
				);
			} else {
				htmlContent.push(`<td style="${styleString}">${cell.formattedValue || ''}</td>`);
			}
		});
		htmlContent.push('</tr>');
	});

	htmlContent.push('</table>');

	return htmlContent.join('\n');
};

// Main function
const xlsx2html = async (fileBytes, sheetName) => {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(fileBytes);

	const worksheet = workbook.getWorksheet(sheetName || 1);
	const data = await worksheetToData(worksheet);
	const htmlTable = renderTable(data);

	return `
		<!DOCTYPE html>
		<html lang="en">
		<head>
			<meta charset="UTF-8">
			<title>Title</title>
		</head>
		<body>
			${htmlTable}
		</body>
		</html>
	`;
};

module.exports = { xlsx2html }; // Export the function
