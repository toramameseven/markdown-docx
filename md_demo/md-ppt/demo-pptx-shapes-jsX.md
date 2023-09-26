<!-- oox param v top h left dpi d-->

# slide title

## shape erro


sss



```js:ppt
const array = [
	{
		type: "text",
		params:[
			[
				{ text: "Sub" },
				{ text: "Subscript", options: { subscript: true } },
				{ text: " // Super" },
				{ text: "Superscript", options: { superscript: true } },
			],
			{ x: 10, y: 6.3, w: 3.3, h: 1.0 }
		]
	},
];
module.exports = { array}; // もしくは module.exports = config
```
sss

```js:ppt

const array = [
	{
		type: "text",
		params:[
			[   { text: "Sub" },
				{ text: "Subscript", options: { subscript: true } },
				{ text: " // Super" },
				{ text: "Superscript", options: { superscript: true } }
			],
			{
				shape: "homePlate",
				x: 2.2,
				y: 0.8,
				w: 3.0,
				h: 1.5,
				fill: { type: "solid", color: "FFdd11", transparency: 50 },
				align: "center",
				fontSize: 14,
			}
		]
	},
];
module.exports = { array}; // もしくは module.exports = config

```

eee

```js:ppt
const array = [
	{
		type: "table",
		params:[
			[
				[
					{ text: "White", options: { fill: { color: "6699CC" }, color: "FFFFFF" } },
					{ text: "Yellow", options: { fill: { color: "99AACC" }, color: "FFFFAA" } },
					{ text: "Pink", options: { fill: { color: "AACCFF" }, color: "E140FE" } },
				],
				[
					{ text: "12pt", options: { fill: { color: "FF0000" }, fontSize: 12 } },
					{ text: "20pt", options: { fill: { color: "00FF00" }, fontSize: 20 } },
					{ text: "28pt", options: { fill: { color: "0000FF" }, fontSize: 28 } },
				],
				[
					{ text: "Bold", options: { fill: { color: "003366" }, bold: true } },
					{ text: "Underline", options: { fill: { color: "336699" }, underline: { style: "sng" } } },
					{ text: "0.15 margin", options: { fill: { color: "6699CC" }, margin: 0.15 } },
				],
			],
			{
				x: 6.0,
				y: 1.1,
				w: 7.0,
				rowH: 0.75,
				fill: { color: "F7F7F7" },
				color: "FFFFFF",
				fontSize: 16,
				valign: "center",
				align: "center",
				border: { pt: "1", color: "FFFFFF" },
			}
		]
	}
];
module.exports = { array }; // もしくは module.exports = config
```