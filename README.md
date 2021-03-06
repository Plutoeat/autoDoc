# autoDoc

## 简单说明

==做这个项目的初衷是为了减少做诸如设置字体样式，页边距等等重复性繁琐的操作。==

当然只用`python-docx`库能实现的功能有限，但我认为已经足够了，其他复杂的操作大可以在`word`上自己修改目前仅做了论文模板。

每个学校要求的论文格式说完全一样也不是，差别很大也不至于，所以一些简单的设置可以直接在`model/config.yaml`文件中修改。

本项目做了三个适应性脚本

- `python-docx`库字体仅支持磅、厘米等不支持字号设置，我简单的对应了小字号和磅值
- `python-docx`不支持首行缩进字符，当然配置太过麻烦就违背了项目的初衷，我默认的是设定首行缩进就缩进两字符，如果有其他需求就请修改代码或直接在`word`中修改了

- 第三个是配置字体对齐方式，但这个做完发现没什么用，因为这都是要添加文本的时候使用，正常我们也不会用这个项目写`word`



另外就是一些说明

- 对字体颜色、字体下划线等有特殊需求的没办法使用该项目生成，因为大家都有不同的需求，如果做太多配置，反而违背了本项目的初衷，所以对`word`有高级定制需求的还请自行修改代码或直接用`word`。本项目仅针对一些简单的重复的需求。我是建议直接用`word`。
- 还有就是论文模板中封面每个学校都是不同的，要做一个适应大多数人的封面模板就配置太过繁琐了，所以建议直接复制粘贴学校模板或者自行设计。
- 论文模板中摘要正文、绪论正文、正文都是采用的`Normal`样式，而结论、参考文献、附录及其正文因为本人学校有特殊要求，则根据要求对`论文标题 1`,`Normal`进行修改后生成的，本质上使用的还是`论文标题 1`和`Normal`，若是将这些的配置也自定义那就太过繁琐还不如直接word配置，本项目目的在于最大限度减少繁琐重复的工作。
- 自行定制样式时，请不要修改我以设置的样式的名字和顺序，其他可以修改或根据下面的模板自行添加。
- 当然若你会修改代码可以自行修改代码以适应你自己的需求，不会需改可以直接在`word`中操作，工作量相对自己完全配置格式会小很多



## 使用方法

命令提示符中

```shell
python main.py your_docx_name type
```

type目前只有论文
