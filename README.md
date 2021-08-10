# Farm-Agricultural-Diagnostics-tool
Macro-enabled workbook for farm agricultural performance diagnostics and visualization, using soil nutrient and water use efficiency related data.

## Table of Contents
- [Introduction](#introduction)
- [Technologies](#technologies)
- [Setup](#setup)
- [Macro code used](#macro-code-used)
- [Acknowledgments](#acknowledgments)

## Introduction
The FAD tool was developed to support agricultural advisory development to become a more demand-driven and data-based process, leading to increased farmer uptake of agricultural advisories. It can be used by academic, government and non-government agencies working in the agricultural sector to run agricultural performance diagnostics based on soil nutrient and water related measurements. It can then categorize farm performance into different classes with an added visual representation. The FAD tool has a straightforward workflow, and is packaged into a user-friendly GUI using Macro-in-Excel VBA. 

## Technologies
- Macro-in-Excel VBA (Microsoft Office Professional Plus 2019, Microsoft Excel)

## Setup
To run the FAD tool: 
1. Download the Excel macro-enabled workbook [FAD-v1.0.xlsm](/FAD-v1.0.xlsm)
2. Enable Macros by 
    - Clicking on "Enable Content" when the "Security Warning (Macros have been disabled.)" appears on the yellow Message bar, or 
    - File tab --> Options --> Trust Center --> Trust Center Settings --> Macro Settings --> "Enable all macros (not recommended, potentially dangerous code can run)."
3. Follow the steps listed in the [FAD-v1.0_Instruction-Manual.pdf](/FAD-v1.0_Instruction-Manual.pdf) file, to run the FAD tool.

## Macro code used
The different macro codes used in the FAD-v1.0 tool are given in [FAD-v1.0_macro-codes](/FAD-v1.0_macro-codes). The nomenclature for each sub and function is identical in the excel workbook and the text file.

## Acknowledgments
The basic macro code used to generate the final visualization chart (Sub Visualization), is attributed to [Tim Williams](https://stackoverflow.com/a/68414093/459088)
