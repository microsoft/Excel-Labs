# Excel Labs, a Microsoft Garage project

Excel Labs is an Office Add-in for Excel that allows us to release experimental Excel features and gather customer feedback about these features. Some Excel Labs ideas may never be incorporated into Excel, but we think the experimentation and feedback this project offers are vital.

Excel Labs currently includes two features:

1. Advanced formula environment: An interface designed for authoring, editing, and reusing formulas.
1. LABS.GENERATIVEAI: A custom function that enables you to send prompts to a generative AI model and return the responses directly to the grid.

The [Microsoft Garage](https://garage.microsoft.com) is an outlet for experimental projects for you to try.

## Advanced formula environment

Advanced formula environment makes it easy to create, edit, and reuse formulas and named LAMBDA functions.

While Excel Name Manager lets you name and reuse any formula, including functions defined with LAMBDA, the interface makes it difficult to author these formulas. Common features that make programming easier are missing, such as immediate inline errors and syntax highlighting. Advanced formula environment fills this gap. It’s an interface for the Name Manager that is designed for formula authoring. Using the advanced formula environment, you can:

- Write formulas using an editor that supports inline errors, IntelliSense, comments, and more.
- Indent formulas, making them easier to read.
- Edit modules of named formulas using a single code editor.
- Quickly reuse LAMBDA formulas by importing them from GitHub gists, or by copying them for other workbooks.

See [Advanced formula environment](/advanced-formula-environment/README.md) for more information.

## LABS.GENERATIVEAI function

This add-in creates a custom function called LABS.GENERATIVEAI that you can use to send prompts from the Excel grid to a generative AI model. LABS.GENERATIVEAI returns the results from the model back to your worksheet.

You can send the model simple or complex prompts, such as requests to:

- Analyze public information.
- Process data and define the format.
- Answer factual or creative questions.
- Produce a response based on a sample.

LABS.GENERATIVEAI allows you to reference other cells in your workbook, and it can be called inside any Excel cell or named formula in the workbook.

See [LABS.GENERATIVEAI function](/labs-generative-ai/README.md) for more information.

## Platform support

Excel Labs works in Excel for Desktop, Mac, and on the web, without installing any additional software. To get started, install the add-in from the Office Store.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.
