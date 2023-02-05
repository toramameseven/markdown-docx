<!-- word export demo-table-merge2.md-->
## merge cells No.2

markdown

```
<!-- word emptyMerge -->

cell(4,2) is not merged. (comment cell)

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |
```
result

<!-- word emptyMerge -->

cell(4,2) is not merged. (comment cell)

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |

## table with emphasis

markdown

```
| data1-1 | data1-2                   |
| ------- | --------------------  --- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |
```

result

| data1-1 | data1-2                   |
| ------- | ------------------------- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |


