# Poi delete and close

## 実行

``` bash
input_path=$(pwd)/src/main/resources/input.xlsx
output_path=$(pwd)/src/main/resources/output.xlsx

mvn clean compile exec:java -Dexec.mainClass="ittimfn.poi.App" -Dexec.args="'$input_path' '$output_path'"
```
