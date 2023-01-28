<!-- word export demo-table-rowmerge.md-->
## Right aligned and rows merge

markdown

```
<!-- word cols 1,3 -->
<!-- word rowMerge 3-4 -->
| data1-1 | data1-2                 |
| -------:| -----------------------:|
| data2-1 | data2-2 XXXXX           |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |
```

result

<!-- word cols 1,3 -->
<!-- word rowMerge 3-4 -->
| data1-1 | data1-2                 |
| -------:| -----------------------:|
| data2-1 | data2-2 XXXXX           |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |

NOTE: aligned is not worked


## table with new line

markdown

```
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<bR>data3-2-2    |
| data4-1 | data4-2                 |
```
result

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<BR>data3-2-2    |
| data4-1 | data4-2                 |

NOTE: does not work



# Links

