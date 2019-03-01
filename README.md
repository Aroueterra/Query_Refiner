[![NPM Version][npm-image]][npm-url]
[![Build Status][travis-image]][travis-url]
[![Downloads Stats][npm-downloads]][npm-url]

---

# Query Refiner

<img align="right" width="100" height="100" src="https://avatars1.githubusercontent.com/u/20365551?s=400&u=e500e44c444dc1edd386184520cef4cbb79c448c&v=4">

> [C#/.Net] This program is used to aggregate information from a spreadsheet and use an internal map to redefine the values of an Income Statement document.

The Query Refiner program is designed to modify and produce an Income Statement/Balance Sheet report. 
It is fed a file containing unorganized values and reorganizes them in a data-value dictionary. After running through an excel document, the program will decide whether it needs to revise the document with the new values based on a criteria dictionary.

At the moment, the dictionary cannot be altered; however, in a future version one will be able to deserialize and serialize the dictionary in and out of the program through JSon use.

The UI is quite clunky, as this was an old project built in the .Net framework. A wpf implementation will help to bring this into the modern era! For insight on a previous WPF project, see: Sage-Pace, a management system.

---

```sh
The program was designed with the specifications and requirements evaluated by the Tax Compliance department of Convergys Philippines Inc., as such it may not be as flexible as I'd intended. Future revisions will rectify this.
```


## Installation

Windows:

```sh
Download file from releases & extract
Run Query_Refiner.exe

*Note: The file and corresponding database is nested in the user's temporary files directory. This is due to the initial requirement that the installation bypass administrator elevation requirements.
```



## Type selection: Income Statement / Balance Sheet

[![Preview](https://github.com/Aroueterra/Query-Refiner/blob/master/graphics/Table.gif)]()


_For more examples and usage, please refer to the [Wiki][wiki]._

## Modifying the Values via the Dictionary

[![Table](https://github.com/Aroueterra/Query-Refiner/blob/master/graphics/Navigating.gif)]()

 
```sh
TBD
```

## Release History


* 0.0.1
    * Initial release

## Meta

August Bryan N. Florese – [@Aroueterra](https://www.facebook.com/Aroueterra) – aroueterra@gmail.com

Distributed under the Mit license. See ``LICENSE`` for more information.

[https://github.com/Aroueterra/](https://github.com/Aroueterra/)

## Contributing

1. Fork it (<https://github.com/yourname/yourproject/fork>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

<!-- Markdown link & img dfn's -->
[npm-image]: https://img.shields.io/npm/v/datadog-metrics.svg?style=flat-square
[npm-url]: https://npmjs.org/package/datadog-metrics
[npm-downloads]: https://img.shields.io/npm/dm/datadog-metrics.svg?style=flat-square
[travis-image]: https://img.shields.io/travis/dbader/node-datadog-metrics/master.svg?style=flat-square
[travis-url]: https://travis-ci.org/dbader/node-datadog-metrics
[wiki]: https://github.com/yourname/yourproject/wiki
