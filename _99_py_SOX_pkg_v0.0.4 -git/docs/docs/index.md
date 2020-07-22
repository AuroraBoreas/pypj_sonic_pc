# Overview

There is a test/measurement for TV Remote Field Sound, which requires precise and time-consuming work to find results.

Goal here is to automate this test / measurement to get accurate result and decrease labor time.

We had explored SOX documentation to get details of each sox command for this automation.

[SoX - Sound eXhange | Documentation](http://sox.sourceforge.net/Docs/Documentation)

## Changelog

- v0.0.1, inital version
- v0.0.2, fix bugs
- v0.0.3, algorithm
  
## Performance

Here is performance comparison between manual operation and this program.

- previous worktime to get results by manual operation<br>it took **2 workdays**
- current worktime using this program<br>it takes **22mins**

    <img src="static/performance.jpg" width="400" />

## Thoughts

- Discuss with Producation Design team to achieve **insights** of this test / measurement.
- Draw a **schema** to optimize and arrange work per step.
- Build **modules** for each work, then debug and test.
- Create a **main** module to combine sub modules, test and run.

## Schema

<img src="static/schema.png" width="800" />

## Confirmation

To make this project maintainable and sustainable, some confirmation as follows.

* Create a module to run **batch** files. `Python builtin package: subprocess, os`
* Build a moulde to **convert bytes** data from 32bit to 16bit. `Python builtin package: array, numpy`
* Create a module to log **sox stats** of test .wav files. `Python builtin package: subprocess`
* Build a module to **extract data** from log file. `Python builtin File I/O`

## Project layout

    +---py_SOX_pkg
        |   main.py
        |   multiple10.py
        |   readme.html
        |   readme.md
        |   readme.pdf
        |   single.py
        |   tree.txt
        |   
        +---data
        |   |   sox_mic_src_stats
        |   |   
        |   +---output
        |   |       aec_loopback_post.bin
        |   |       sox_stats_diff
        |   |       
        |   +---已转
        |   |       aec_loopback_post.bin
        |   |       
        |   \---未转
        |       |   aec_loopback_post.bin
        |       |   
        |       \---bin
        |               aec_loopback_post.bin
        |               
        +---docs
        |   |   mkdocs.yml
        |   |   
        |   \---docs
        |       |   about.md
        |       |   algorithm.md
        |       |   detail.md
        |       |   index.md
        |       |   
        |       \---static
        |               IDEL_LOG.png
        |               log.txt
        |               mic_rms_src_rms.png
        |               mic_rms_variant.png
        |               performance.jpg
        |               schema.png
        |               
        +---lib
        |   |   core.py
        |   |   
        |   +---pkg
        |   |       converter.py
        |   |       gain.dll
        |   |       gain.py
        |   |       parser.py
        |   |       runbat.py
        |   |       stats.py
        |   |       __init__.py
        |   |       
        |   \---util
        |           hashes.py
        |           
        \---tests
            |   Ref_M5_MicAuto_SRC_Adjust8.py
            |   test_converter.py
            |   test_gain.py
            |   test_parser.py
            |   test_runbat.py
            |   
            \---bats
                    STEP1 create_files.cmd
                    STEP2 rename_files.cmd
                    STEP3 delete_files.cmd