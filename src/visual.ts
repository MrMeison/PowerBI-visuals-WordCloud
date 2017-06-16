/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {
    // d3
    import Selection = d3.Selection;
    import UpdateSelection = d3.selection.Update;
    import Transition = d3.Transition;



    // powerbi
    import IViewport = powerbi.IViewport;
    import DataView = powerbi.DataView;
    import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
    import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
    import VisualObjectInstance = powerbi.VisualObjectInstance;
    import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;

    // powerbi.extensibility
    import IVisual = powerbi.extensibility.IVisual;
    import IColorPalette = powerbi.extensibility.IColorPalette;

    // powerbi.extensibility.visual
    import IVisualHost = powerbi.extensibility.visual.IVisualHost;
    import ISelectionIdBuilder = powerbi.extensibility.ISelectionIdBuilder;
    import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
    import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;

    // powerbi.visuals
    import ISelectionId = powerbi.visuals.ISelectionId;

    // powerbi.extensibility.utils.svg
    import IMargin = powerbi.extensibility.utils.svg.IMargin;
    import translate = powerbi.extensibility.utils.svg.translate;
    import IPoint = powerbi.extensibility.utils.svg.shapes.IPoint;
    import translateAndScale = powerbi.extensibility.utils.svg.translateAndScale;
    import ClassAndSelector = powerbi.extensibility.utils.svg.CssConstants.ClassAndSelector;
    import createClassAndSelector = powerbi.extensibility.utils.svg.CssConstants.createClassAndSelector;

    // powerbi.extensibility.utils.type
    import PixelConverter = powerbi.extensibility.utils.type.PixelConverter;

    // powerbi.extensibility.utils.formatting
    import ValueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;
    import IValueFormatter = powerbi.extensibility.utils.formatting.IValueFormatter;
    import textMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;

    // powerbi.extensibility.utils.color
    import ColorHelper = powerbi.extensibility.utils.color.ColorHelper;

    enum WordCloudScaleType {
        logn,
        sqrt,
        value
    }

    export class WordCloud implements IVisual {
        private static ClassName: string = "wordCloud";

        private static Words: ClassAndSelector = createClassAndSelector("words");
        private static VisualInteractive: ClassAndSelector = createClassAndSelector("wordCloud--interactive");
        private static WordGroup: ClassAndSelector = createClassAndSelector("word");
        private static SelectedWord: ClassAndSelector = createClassAndSelector("word--selected");

        private static StopWordsDelimiter: string = " ";

        // private static Radians: number = Math.PI / 180;

        private static MinOpacity: number = 0.2;
        private static MaxOpacity: number = 1;

        private static Punctuation: string[] = [
            "!", ".", ":", "'", ";", ",", "?",
            "@", "#", "$", "%", "^", "&", "*",
            "(", ")", "[", "]", "\"", "\\", "/",
            "-", "_", "+", "=", "<", ">", "|"
        ];

        private static StopWords: string[] = [
            "a", "able", "about", "across", "after", "all", "almost", "also", "am", "among", "an",
            "and", "any", "are", "as", "at", "be", "because", "been", "but", "by", "can", "cannot",
            "could", "did", "do", "does", "either", "else", "ever", "every", "for", "from", "get",
            "got", "had", "has", "have", "he", "her", "hers", "him", "his", "how", "however", "i",
            "if", "in", "into", "is", "it", "its", "just", "least", "let", "like", "likely", "may",
            "me", "might", "most", "must", "my", "neither", "no", "nor", "not", "of", "off", "often",
            "on", "only", "or", "other", "our", "own", "rather", "said", "say", "says", "she", "should",
            "since", "so", "some", "than", "that", "the", "their", "them", "then", "there", "these",
            "they", "this", "tis", "to", "too", "twas", "us", "wants", "was", "we", "were", "what",
            "when", "where", "which", "while", "who", "whom", "why", "will", "with", "would", "yet",
            "you", "your"
        ];

        private static DefaultMargin: IMargin = {
            top: 10,
            right: 10,
            bottom: 10,
            left: 10
        };

        private static MinViewport: IViewport = {
            width: 0,
            height: 0
        };

        private static DataPointFillProperty: DataViewObjectPropertyIdentifier = {
            objectName: "dataPoint",
            propertyName: "fill"
        };

        /**
         * Names of these consts aren't good, but I have no idea how to call them better.
         * TODO: Please rename them if you know any better names.
         */
        // private static TheFirstLineHeight: string = PixelConverter.toString(5); // Note: This construction fixes bug #6343.
        // private static TheSecondLineHeight: string = PixelConverter.toString(10); // Note: This construction fixes bug #6343.
        private static TheThirdLineHeight: string = PixelConverter.toString(14); // Note: This construction fixes bug #6343.
        private static TheFourthLineHeight: string = PixelConverter.toString(15); // Note: This construction fixes bug #6343.
        private static MinCount: number = 1;
        private static DefaultX: number = 0;
        private static DefaultY: number = 0;
        private static DefaultPadding: number = 1;
        private static DefaultWidth: number = 0;
        private static DefaultHeight: number = 0;
        private static DefaultXOff: number = 0;
        private static DefaultYOff: number = 0;
        private static DefaultX0: number = 0;
        private static DefaultY0: number = 0;
        private static DefaultX1: number = 0;
        private static DefaultY1: number = 0;
        private static MinFontSize: number = 0;
        private static DefaultAngle: number = 0;
        private static FontSizePercentage: number = 100;
        private get settings(): WordCloudSettings {
            return this.data && this.data.settings;
        }

        private data: WordCloudData;
        private colorPalette: IColorPalette;
        private durationAnimations: number = 50;
        private root: Selection<any>;
        private main: Selection<any>;
        private wordsContainerSelection: Selection<any>;
        private wordsTextUpdateSelection: UpdateSelection<WordCloudDataPoint>;
        private layout: VisualLayout;
        private d3Layout: d3.layout.Cloud<d3.layout.cloud.Word>;
        private visualHost: IVisualHost;
        private selectionManager: ValueSelectionManager<string>;
        private visualUpdateOptions: VisualUpdateOptions;
        private isUpdating: boolean = false;
        private incomingUpdateOptions: VisualUpdateOptions;
        private oldIdentityKeys: string[];
        public static converter(
            dataView: DataView,
            colors: IColorPalette,
            visualHost: IVisualHost,
            previousData: WordCloudData): WordCloudData {

            let categorical: WordCloudColumns<DataViewCategoryColumn>,
                catValues: WordCloudColumns<any[]>,
                settings: WordCloudSettings,
                colorHelper: ColorHelper,
                stopWords: string[],
                texts: WordCloudText[] = [],
                reducedTexts: WordCloudText[][],
                dataPoints: WordCloudDataPoint[],
                wordValueFormatter: IValueFormatter,
                queryName: string;

            categorical = WordCloudColumns.getCategoricalColumns(dataView);

            if (!categorical || !categorical.Category || _.isEmpty(categorical.Category.values)) {
                return null;
            }

            catValues = WordCloudColumns.getCategoricalValues(dataView);
            settings = WordCloud.parseSettings(dataView, previousData && previousData.settings);

            wordValueFormatter = ValueFormatter.create({
                format: ValueFormatter.getFormatStringByColumn(categorical.Category.source)
            });

            stopWords = !!settings.stopWords.words && _.isString(settings.stopWords.words)
                ? settings.stopWords.words.split(WordCloud.StopWordsDelimiter)
                : [];

            stopWords = settings.stopWords.isDefaultStopWords
                ? stopWords.concat(WordCloud.StopWords)
                : stopWords;

            colorHelper = new ColorHelper(
                colors,
                WordCloud.DataPointFillProperty,
                wordCloudUtils.getRandomColor());

            queryName = (categorical.Values
                && categorical.Values[0]
                && categorical.Values[0].source
                && categorical.Values[0].source.queryName)
                || null;
            for (let index: number = 0; index < catValues.Category.length; index += 1) {
                let item: any = catValues.Category[index];
                if (!item) {
                    continue;
                }
                let color: string;
                let selectionIdBuilder: ISelectionIdBuilder;
                if (categorical.Category.objects && categorical.Category.objects[index]) {
                    color = colorHelper.getColorForMeasure(categorical.Category.objects[index], "");
                } else {
                    color = previousData && previousData.texts && previousData.texts[index]
                        ? previousData.texts[index].color
                        : colors.getColor(index.toString()).value;
                }

                selectionIdBuilder = visualHost.createSelectionIdBuilder()
                    .withCategory(dataView.categorical.categories[0], index);

                if (queryName) {
                    selectionIdBuilder.withMeasure(queryName);
                }

                item = wordValueFormatter.format(item);
                texts.push({
                    text: item,
                    count: (catValues.Values
                        && catValues.Values[index]
                        && !isNaN(catValues.Values[index]))
                        ? catValues.Values[index]
                        : WordCloud.MinCount,
                    index: index,
                    selectionId: selectionIdBuilder.createSelectionId() as ISelectionId,
                    color: color,
                    textGroup: item
                });
            }

            reducedTexts = WordCloud.getReducedText(texts, stopWords, settings);
            dataPoints = WordCloud.getDataPoints(reducedTexts, settings);

            return {
                dataView: dataView,
                settings: settings,
                texts: texts,
                dataPoints: dataPoints
            };
        }

        private static parseSettings(dataView: DataView, previousSettings: WordCloudSettings): WordCloudSettings {
            const settings: WordCloudSettings = WordCloudSettings.parse<WordCloudSettings>(dataView);

            settings.general.minFontSize = Math.max(
                settings.general.minFontSize,
                GeneralSettings.MinFontSize);

            settings.general.maxFontSize = Math.max(
                settings.general.maxFontSize,
                GeneralSettings.MinFontSize);

            settings.general.maxFontSize = Math.max(
                settings.general.maxFontSize,
                settings.general.minFontSize);

            settings.rotateText.minAngle = Math.max(
                Math.min(settings.rotateText.minAngle, RotateTextSettings.MaxAngle),
                RotateTextSettings.MinAngle);

            settings.rotateText.maxAngle = Math.max(
                Math.min(settings.rotateText.maxAngle, RotateTextSettings.MaxAngle),
                RotateTextSettings.MinAngle);

            settings.rotateText.maxAngle = Math.max(
                settings.rotateText.maxAngle,
                settings.rotateText.minAngle);

            settings.general.maxNumberOfWords = Math.max(
                Math.min(settings.general.maxNumberOfWords, RotateTextSettings.MaxNumberOfWords),
                RotateTextSettings.MinNumberOfWords);

            settings.rotateText.maxNumberOfOrientations = Math.max(
                Math.min(settings.rotateText.maxNumberOfOrientations, RotateTextSettings.MaxNumberOfWords),
                RotateTextSettings.MinNumberOfWords);

            return settings;
        }

        private static getReducedText(
            texts: WordCloudText[],
            stopWords: string[],
            settings: WordCloudSettings): WordCloudText[][] {

            let brokenStrings: WordCloudText[] = WordCloud.processText(texts, stopWords, settings),
                result: WordCloudText[][] = <WordCloudText[][]>_.values(_.groupBy(
                    brokenStrings,
                    (textObject: WordCloudText) => textObject.text.toLocaleLowerCase()));

            result = result.map((texts: WordCloudText[]) => {
                return _.sortBy(texts, (textObject: WordCloudText) => textObject.textGroup.length);
            });

            return result;
        }

        private static processText(
            words: WordCloudText[],
            stopWords: string[],
            settings: WordCloudSettings): WordCloudText[] {
            let processedText: WordCloudText[] = [],
                partOfProcessedText: WordCloudText[] = [],
                whiteSpaceRegExp: RegExp = /\s/,
                punctuationRegExp: RegExp = new RegExp(`[${WordCloud.Punctuation.join("\\")}]`, "gim");

            words.forEach((item: WordCloudText) => {
                if (typeof item.text === "string") {
                    let splittedWords: string[] = item.text
                        .replace(punctuationRegExp, " ")
                        .split(whiteSpaceRegExp);
                    const splittedWordsOriginalLength: number = splittedWords.length;

                    splittedWords = WordCloud.getFilteredWords(splittedWords, stopWords, settings);
                    partOfProcessedText = settings.general.isBrokenText
                        ? WordCloud.getBrokenWords(splittedWords, item, whiteSpaceRegExp)
                        : WordCloud.getFilteredSentences(splittedWords, item, splittedWordsOriginalLength, settings, punctuationRegExp);

                    processedText.push(...partOfProcessedText);
                } else {
                    processedText.push(item);
                }
            });

            return processedText;
        }

        private static getBrokenWords(
            splittedWords: string[],
            item: WordCloudText,
            whiteSpaceRegExp: RegExp): WordCloudText[] {

            let brokenStrings: WordCloudText[] = [];

            splittedWords.forEach((splittedWord: string) => {
                if (splittedWord.length === 0 || whiteSpaceRegExp.test(splittedWord)) {
                    return;
                }

                brokenStrings.push({
                    text: splittedWord,
                    textGroup: item.textGroup,
                    count: item.count,
                    index: item.index,
                    selectionId: item.selectionId,
                    color: item.color
                });
            });

            return brokenStrings;
        }

        private static getFilteredSentences(
            splittedWords: string[],
            item: WordCloudText,
            splittedWordsOriginalLength: number,
            settings: WordCloudSettings,
            punctuationRegExp: RegExp): WordCloudText[] {

            if (!settings.general.isPunctuationsCharacters) {
                item.text = item.text
                    .replace(punctuationRegExp, " ");
            }

            if (splittedWords.length === splittedWordsOriginalLength) {
                return [item];
            }

            return [];
        }

        private static getFilteredWords(
            words: string[],
            stopWords: string[],
            settings: WordCloudSettings) {

            if (!settings.stopWords.show || !stopWords.length) {
                return words;
            }

            return words.filter((value: string) => {
                return value.length > 0 && !stopWords.some((removeWord: string) => {
                    return value.toLocaleLowerCase() === removeWord.toLocaleLowerCase();
                });
            });
        }

        private static getDataPoints(
            textGroups: WordCloudText[][],
            settings: WordCloudSettings): WordCloudDataPoint[] {

            if (_.isEmpty(textGroups)) {
                return [];
            }

            const returnValues: WordCloudDataPoint[] = textGroups.map((values: WordCloudText[]) => {
                return {
                    text: values[0].text,
                    x: WordCloud.DefaultX,
                    y: WordCloud.DefaultY,
                    rotate: WordCloud.getAngle(settings),
                    padding: WordCloud.DefaultPadding,
                    width: WordCloud.DefaultWidth,
                    height: WordCloud.DefaultHeight,
                    xOff: WordCloud.DefaultXOff,
                    yOff: WordCloud.DefaultYOff,
                    x0: WordCloud.DefaultX0,
                    y0: WordCloud.DefaultY0,
                    x1: WordCloud.DefaultX1,
                    y1: WordCloud.DefaultY1,
                    color: values[0].color,
                    selectionIds: values.map((text: WordCloudText) => text.selectionId),
                    wordIndex: values[0].index,
                    count: _.sumBy(values, (text: WordCloudText) => text.count)
                } as WordCloudDataPoint;
            });

            const minValue: number = _.minBy(returnValues, (dataPoint: WordCloudDataPoint) => dataPoint.count).count,
                maxValue: number = _.maxBy(returnValues, (dataPoint: WordCloudDataPoint) => dataPoint.count).count,
                texts: WordCloudText[] = textGroups.map((textGroup: WordCloudText[]) => textGroup[0]);

            returnValues.forEach((dataPoint: WordCloudDataPoint) => {
                dataPoint.size = WordCloud.getWordFontSize(
                    texts,
                    settings,
                    dataPoint.count,
                    minValue,
                    maxValue);
            });

            return returnValues.sort((firstDataPoint: WordCloudDataPoint, secondDataPoint: WordCloudDataPoint) => {
                return secondDataPoint.count - firstDataPoint.count;
            });
        }

        private static getWordFontSize(
            texts: WordCloudText[],
            settings: WordCloudSettings,
            value: number,
            minValue: number,
            maxValue: number,
            scaleType: WordCloudScaleType = WordCloudScaleType.value) {

            let weight: number,
                fontSize: number,
                minFontSize: number = settings.general.minFontSize * GeneralSettings.FontSizePercentageFactor,
                maxFontSize: number = settings.general.maxFontSize * GeneralSettings.FontSizePercentageFactor;

            if (texts.length <= RotateTextSettings.MinNumberOfWords) {
                return maxFontSize;
            }

            weight = WordCloud.getWeightByScaleType(value, scaleType);

            if (weight > minValue) {
                fontSize = (maxValue - minValue) !== WordCloud.MinFontSize
                    ? (maxFontSize * (weight - minValue)) / (maxValue - minValue)
                    : WordCloud.MinFontSize;
            } else {
                fontSize = WordCloud.MinFontSize;
            }

            fontSize = (fontSize * WordCloud.FontSizePercentage) / maxFontSize;

            fontSize = (fontSize * (maxFontSize - minFontSize)) / WordCloud.FontSizePercentage + minFontSize;

            return fontSize;
        }

        private static getWeightByScaleType(
            value: number,
            scaleType: WordCloudScaleType = WordCloudScaleType.value): number {

            switch (scaleType) {
                case WordCloudScaleType.logn: {
                    return Math.log(value);
                }
                case WordCloudScaleType.sqrt: {
                    return Math.sqrt(value);
                }
                case WordCloudScaleType.value:
                default: {
                    return value;
                }
            }
        }

        private static getAngle(settings: WordCloudSettings): number {
            if (!settings.rotateText.show) {
                return WordCloud.DefaultAngle;
            }

            const angle: number = ((settings.rotateText.maxAngle - settings.rotateText.minAngle)
                / settings.rotateText.maxNumberOfOrientations)
                * Math.floor(Math.random() * settings.rotateText.maxNumberOfOrientations);

            return settings.rotateText.minAngle + angle;
        }

        constructor(options: VisualConstructorOptions) {
            this.init(options);
        }

        public init(options: VisualConstructorOptions): void {
            this.root = d3.select(options.element).append("svg");

            this.colorPalette = options.host.colorPalette;
            this.visualHost = options.host;

            this.selectionManager = new ValueSelectionManager<string>(
                this.visualHost,
                (text: string): ISelectionId[] => {
                    const dataPoints: WordCloudDataPoint[] = this.data
                        && this.data.dataPoints
                        && this.data.dataPoints.filter((dataPoint: WordCloudDataPoint) => {
                            return dataPoint.text === text;
                        });

                    return dataPoints && dataPoints[0] && dataPoints[0].selectionIds
                        ? dataPoints[0].selectionIds
                        : [];
                });

            this.layout = new VisualLayout(null, WordCloud.DefaultMargin);

            this.root
                .classed(WordCloud.ClassName, true)
                .attr("width", "100%")
                .attr("height", "100%");

            this.root.on("click", () => {
                this.clearSelection();
            });

            this.wordsContainerSelection = this.root
                .append("g")
                .classed(WordCloud.Words.className, true);
            this.d3Layout = d3.layout.cloud()
                .padding(5)
                .rotate(() => ~~(Math.random() * 2) * Math.random() * 90)
                .font(this.root.style("font-family"))
                .fontSize((d: WordCloudDataPoint) => d.size)
                .text((d: WordCloudDataPoint) => d.text)
                .on("end", this.render.bind(this));
        }

        public update(visualUpdateOptions: VisualUpdateOptions): void {
            if (!visualUpdateOptions
                || !visualUpdateOptions.viewport
                || !visualUpdateOptions.dataViews
                || !visualUpdateOptions.dataViews[0]
                || !visualUpdateOptions.viewport
                || !(visualUpdateOptions.viewport.height >= WordCloud.MinViewport.height)
                || !(visualUpdateOptions.viewport.width >= WordCloud.MinViewport.width)) {
                return;
            }
            if (visualUpdateOptions !== this.visualUpdateOptions) {
                this.incomingUpdateOptions = visualUpdateOptions;
            }


            if (!this.isUpdating && (this.incomingUpdateOptions !== this.visualUpdateOptions)) {
                this.visualUpdateOptions = this.incomingUpdateOptions;
                this.layout.viewport = this.visualUpdateOptions.viewport;

                const dataView: DataView = visualUpdateOptions.dataViews[0];

                if (this.layout.viewportInIsZero) {
                    return;
                }
                this.updateSize(visualUpdateOptions.viewport);
                this.data = WordCloud.converter(dataView, this.colorPalette, this.visualHost, this.data);

                if (!this.data) {
                    this.clear();
                    return;
                }

                this.updateWordSet(this.data.dataPoints);
                this.d3Layout.start();
            }
        }

        private updateSize(viewport: IViewport) {
            this.d3Layout.size([viewport.width, viewport.height]);
        }

        private updateWordSet(points: WordCloudDataPoint[]): void {
            this.d3Layout.words(points);
        }

        private render(points: WordCloudDataPoint[]): void {
            if (!points || !points.length) {
                return;
            }
            this.clear();
            this.wordsContainerSelection
                .attr("width", this.d3Layout.size()[0])
                .attr("height", this.d3Layout.size()[1])
                .attr("transform", "translate(" + this.d3Layout.size()[0] / 2 + "," + this.d3Layout.size()[1] / 2 + ")")
                .selectAll("text")
                .data(points)
                .enter()
                .append("text")
                .classed(WordCloud.WordGroup.className, true)
                .style({
                    "font-size": ((item: WordCloudDataPoint): string => PixelConverter.toString(item.size)),
                    "fill": ((item: WordCloudDataPoint): string => item.color),
                })
                .attr("transform", (d: WordCloudDataPoint) => "translate(" + [d.x, d.y] + ")rotate(" + d.rotate + ")")
                .text((dataPoint: WordCloudDataPoint) => dataPoint.text)
                .on("click", (dataPoint: WordCloudDataPoint) => {
                    const event = (d3.event as MouseEvent);
                    event.stopPropagation();
                    d3.select(event.target).classed(WordCloud.SelectedWord.className);
                    this.root.classed(WordCloud.VisualInteractive.className, true);
                    this.setSelection(dataPoint);
                });
        }
        private clear(): void {
            this.root
                .select(WordCloud.Words.selectorName)
                .selectAll(WordCloud.WordGroup.selectorName)
                .remove();
        }

        private clearIncorrectSelection(dataView: DataView): void {
            let categories: DataViewCategoryColumn[],
                identityKeys: string[],
                oldIdentityKeys: string[] = this.oldIdentityKeys;

            categories = dataView
                && dataView.categorical
                && dataView.categorical.categories;

            identityKeys = categories
                && categories[0]
                && categories[0].identity
                && categories[0].identity.map((identity: DataViewScopeIdentity) => identity.key);

            this.oldIdentityKeys = identityKeys;

            if (oldIdentityKeys && oldIdentityKeys.length > identityKeys.length) {
                this.selectionManager.clear(false);

                return;
            }

            if (!_.isEmpty(identityKeys)) {
                let incorrectValues: SelectionIdValues<string>[] = this.selectionManager
                    .getSelectionIdValues
                    .filter((idValue: SelectionIdValues<string>) => {
                        return idValue.selectionId.some((selectionId: ISelectionId) => {
                            return _.includes(identityKeys, selectionId.getKey());
                        });
                    });

                incorrectValues.forEach((value: SelectionIdValues<string>) => {
                    this.selectionManager
                        .selectedValues
                        .splice(this.selectionManager
                            .selectedValues
                            .indexOf(value.value), 1);
                });
            }
        }

        private setSelection(dataPoint: WordCloudDataPoint): void {
            if (!dataPoint) {
                this.clearSelection();

                return;
            }

            this.selectionManager
                .selectAndSendSelection(dataPoint.text, (d3.event as MouseEvent).ctrlKey)
                .then(() => this.renderSelection());
        }

        private clearSelection(): void {
            this.selectionManager
                .clear(true);
            this.root
                .selectAll(WordCloud.SelectedWord.selectorName)
                .classed(WordCloud.SelectedWord.className, false);
            this.root.classed(WordCloud.VisualInteractive.className, false);
        }

        private renderSelection(): void {
            if (!this.wordsTextUpdateSelection) {
                return;
            }

            if (!this.selectionManager.hasSelection) {
                this.setOpacity(this.wordsTextUpdateSelection, WordCloud.MaxOpacity);

                return;
            }

            const selectedColumns: UpdateSelection<WordCloudDataPoint> = this.wordsTextUpdateSelection
                .filter((dataPoint: WordCloudDataPoint) => {
                    return this.selectionManager.isSelected(dataPoint.text);
                });

            this.setOpacity(this.wordsTextUpdateSelection, WordCloud.MinOpacity);
            this.setOpacity(selectedColumns, WordCloud.MaxOpacity);
        }

        private setOpacity(element: Selection<any>, opacityValue: number): void {
            element.style("fill-opacity", opacityValue);

            if (this.main) { // Note: This construction fixes bug #6343.
                this.main.style("line-height", WordCloud.TheThirdLineHeight);

                this.animateSelection(this.main, 0, this.durationAnimations)
                    .style("line-height", WordCloud.TheFourthLineHeight);
            }
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            const settings: WordCloudSettings = this.settings
                ? this.settings
                : WordCloudSettings.getDefault() as WordCloudSettings;

            let instanceEnumeration: VisualObjectInstanceEnumeration =
                WordCloudSettings.enumerateObjectInstances(settings, options);

            switch (options.objectName) {
                case "dataPoint": {
                    if (this.data && this.data.dataPoints) {
                        this.enumerateDataPoint(options, instanceEnumeration);
                    }

                    break;
                }
            }

            return instanceEnumeration || [];
        }

        private enumerateDataPoint(
            options: EnumerateVisualObjectInstancesOptions,
            instanceEnumeration: VisualObjectInstanceEnumeration): void {

            let uniqueDataPoints: WordCloudDataPoint[] = _.uniqBy(
                this.data.dataPoints,
                (dataPoint: WordCloudDataPoint) => dataPoint.wordIndex);

            this.enumerateDataPointColor(uniqueDataPoints, options, instanceEnumeration);
        }

        private enumerateDataPointColor(
            dataPoints: WordCloudDataPoint[],
            options: EnumerateVisualObjectInstancesOptions,
            instanceEnumeration: VisualObjectInstanceEnumeration): void {

            let wordCategoriesIndex: number[] = [];

            dataPoints.forEach((item: WordCloudDataPoint) => {
                if (wordCategoriesIndex.indexOf(item.wordIndex) === -1) {
                    let instance: VisualObjectInstance;

                    wordCategoriesIndex.push(item.wordIndex);

                    instance = {
                        objectName: options.objectName,
                        displayName: this.data.texts[item.wordIndex].text,
                        selector: ColorHelper.normalizeSelector(
                            item.selectionIds[0].getSelector(),
                            false),
                        properties: { fill: { solid: { color: item.color } } }
                    };

                    this.addAnInstanceToEnumeration(instanceEnumeration, instance);
                }
            });

        }

        private addAnInstanceToEnumeration(
            instanceEnumeration: VisualObjectInstanceEnumeration,
            instance: VisualObjectInstance): void {

            if ((instanceEnumeration as VisualObjectInstanceEnumerationObject).instances) {
                (instanceEnumeration as VisualObjectInstanceEnumerationObject)
                    .instances
                    .push(instance);
            } else {
                (instanceEnumeration as VisualObjectInstance[]).push(instance);
            }
        }

        private animateSelection<T extends Selection<any>>(
            element: T,
            duration: number = 0,
            delay: number = 0,
            callback?: (data: any, index: number) => void): Transition<any> {

            return element
                .transition()
                .delay(delay)
                .duration(duration)
                .each("end", callback);
        }

        public destroy(): void {
            this.root = null;
        }
    }
}
