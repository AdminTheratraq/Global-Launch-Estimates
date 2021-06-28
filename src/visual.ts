/*
*  Power BI Visual CLI
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
"use strict";
import "regenerator-runtime/runtime";
import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataViewObjects = powerbi.DataViewObjects;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import * as d3 from 'd3';
import { VisualSettings } from "./settings";
import Datamap from 'datamaps';
import * as sanitizeHtml from 'sanitize-html';
import * as validDataUrl from 'valid-data-url';

export interface GlobalFacilityLocation {
    Company: string;
    Region: string;
    State: string;
    Country: string;
    DocumentLink: string;
    Launch: string;
    Color: string;
    Highlights: string;
    HeaderImage: string;
    FooterImage: string;
    selectionId: powerbi.visuals.ISelectionId;
}

export interface GlobalFacilityLocations {
    GlobalFacilityLocation: GlobalFacilityLocation[]
}

export function logExceptions(): MethodDecorator {
    return (target: Object, propertyKey: string, descriptor: TypedPropertyDescriptor<any>)
        : TypedPropertyDescriptor<any> => {

        return {
            value: function () {
                try {
                    return descriptor.value.apply(this, arguments);
                } catch (e) {
                    throw e;
                }
            }
        }
    }
}

export function getCategoricalObjectValue<T>(objects: DataViewObjects, index: number, objectName: string, propertyName: string, defaultValue: T): T {
    if (objects) {
        let object = objects[objectName];
        if (object) {
            let property: T = <T>object[propertyName];
            if (property !== undefined) {
                return property;
            }
        }
    }
    return defaultValue;
}

export class Visual implements IVisual {
    private target: d3.Selection<HTMLElement, any, any, any>;
    private header: d3.Selection<HTMLElement, any, any, any>;
    private footer: d3.Selection<HTMLElement, any, any, any>;
    private container: HTMLElement;
    private margin = { top: 50, right: 40, bottom: 50, left: 40 };
    private settings: VisualSettings;
    private host: IVisualHost;
    private map: any;
    private yearColorData: any;
    private events: IVisualEventService;
    private headerImgHeight = 0;
    private footerImgHeight = 0;
    private mainContentHeight = 0;
    private selectionManager: ISelectionManager;

    constructor(options: VisualConstructorOptions) {
        this.header = d3.select(options.element).append('div');
        this.target = d3.select(options.element).append('div');
        this.footer = d3.select(options.element).append('div');
        this.host = options.host;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
    }

    public update(options: VisualUpdateOptions) {
        this.events.renderingStarted(options);
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        this.target.selectAll('*').remove();
        let _this = this;
        this.target.attr('class', 'map-container');
        this.target.attr('style', 'height:' + options.viewport.height + 'px;width:' + options.viewport.width + 'px');
        let gHeight = options.viewport.height - this.margin.top - this.margin.bottom;
        let gWidth = options.viewport.width - this.margin.left - this.margin.right;

        let mapData = Visual.CONVERTER(options.dataViews[0], this.host);
        mapData = mapData.slice(0, 1000);
        let finalMapData: GlobalFacilityLocation[]=[];
        let currentDate, previousyear:number, futureyear:number;
        currentDate = new Date();
        previousyear = currentDate.getFullYear() - 1;
        futureyear = currentDate.getFullYear() + 8;

        for (let i=previousyear; i<=futureyear;i++) {
            mapData.filter((data) => { if (data.Launch === i.toString()) { finalMapData.push(data) } });
        }
        this.target.on("contextmenu", () => {
            const mouseEvent: MouseEvent = <MouseEvent> d3.event;
            const eventTarget: any = mouseEvent.target;
            let dataPoint: any = d3.select(eventTarget).datum();
            let selectionId = null;
            if (dataPoint) {
                selectionId = this.getSelectionId(finalMapData, dataPoint?.Company ? dataPoint.selectionId : dataPoint?.properties?.name);
                if (!selectionId) {
                    selectionId = this.getSelectionId(finalMapData, dataPoint?.Company ? dataPoint.selectionId : dataPoint?.id);
                }
            }
            this.selectionManager.showContextMenu(
              dataPoint ? selectionId : {},
              {
                x: mouseEvent.clientX,
                y: mouseEvent.clientY,
              }
            );
            mouseEvent.preventDefault();
          });

        this.renderHeaderAndFooter(finalMapData, options);

        let countryCodes = this.getCountryCodes();

        let mainContent = this.target.append('div')
            .attr('class', 'main-content');
        
        this.mainContentHeight = options.viewport.height;
        if  (this.settings.locations.headerImage && this.settings.locations.footerImage) { this.mainContentHeight = options.viewport.height - 210; } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) { this.mainContentHeight = options.viewport.height - 105; }
        
        mainContent.style('position', 'relative')
            .style('height', (this.mainContentHeight) + 'px')
            .style('width', (options.viewport.width) + 'px');

        let header = mainContent.append('div')
            .attr('class', 'header');

        if (!this.settings.locations.viewRegionalMap) {
            this.renderWorldMapTitle(header);

            let mapWidth = options.viewport.width - 266;
            if (this.settings.locations.headerImage && this.settings.locations.footerImage) { mapWidth = 620; } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) { mapWidth = 500;}

            let container = mainContent.append('div')
                .attr('class', 'container')
                .style('position', 'relative')
                .style('height', (options.viewport.height - 60.13) + 'px')
                .style('width', (mapWidth) + 'px');

            this.renderWorldMap(finalMapData, countryCodes, container);

            this.renderWorldMapLegend(mainContent, options);
        }
        else {
            
            let currentRegion = this.getCurrentRegion(finalMapData);

            this.renderRegionalMapTitle(header, currentRegion);

            if (currentRegion != "USA") {
                finalMapData = finalMapData.filter(d => d.Region === currentRegion);
            }

            let containerWrap = mainContent.append('div')
                .attr('class', 'container-wrap');

            let regionMapWrap = this.createRegionMapWrapElement(containerWrap, options);

            let regionMap = this.createRegionMapElement(regionMapWrap, options);

            this.renderRegionalMap(finalMapData, countryCodes, regionMap, currentRegion);

            this.createHighlightsContainerElement(containerWrap, finalMapData);

            this.renderRegionalMapLegend(mainContent, currentRegion);
        }
        this.events.renderingFinished(options);
    }

    private renderHeaderAndFooter(mapData: GlobalFacilityLocation[], options) {
        let viewportHeight = options.viewport.height;
        let viewportwidth = options.viewport.width;
        let layoutContentHeight = 0;

        let [globalFacility] = mapData;
        // sanitized user input from settings
        if (this.settings.locations.headerImage) {
            let headerImage = new Image();
            headerImage.onload = () => {
                this.headerImgHeight = headerImage.height;
                layoutContentHeight += headerImage.height;
                this.target.attr('style', 'height:' + (viewportHeight - layoutContentHeight) + 'px;width:' + (viewportwidth) + 'px');
                // removed .html() method and built DOM using append method
                this.header
                    .attr('class', 'visual-header')
                    .attr('style', 'height:' + this.headerImgHeight + 'px;')
                    .append('img')
                    .attr('src', validDataUrl(globalFacility.HeaderImage) ? globalFacility.HeaderImage : '');
            }
            if (validDataUrl(globalFacility.HeaderImage)) {
                headerImage.src = globalFacility.HeaderImage;
            }
        }
        else {
            this.header.selectAll('img').remove();
            this.header.classed('visual-header', false);
            this.header.style('height', 'auto');
        }

        // sanitized user input from settings
        if (this.settings.locations.footerImage) {
            let footerImage = new Image();
            footerImage.onload = () => {
                this.footerImgHeight = footerImage.height;
                layoutContentHeight += footerImage.height;
                this.target.attr('style', 'height:' + (viewportHeight - layoutContentHeight) + 'px;width:' + (viewportwidth) + 'px');
                // removed .html() method and built DOM using append method
                this.footer
                    .attr('class', 'visual-footer')
                    .attr('style', 'height:' + this.footerImgHeight + 'px;')
                    .append('img')
                    .attr('src', validDataUrl(globalFacility.FooterImage) ? globalFacility.FooterImage : '');
            }
            if (validDataUrl(globalFacility.FooterImage)) {
                footerImage.src = globalFacility.FooterImage;
            }
        }
        else {
            this.footer.selectAll('img').remove();
            this.footer.classed('visual-footer', false);
            this.footer.style('height', 'auto');
        }

    }

    private renderWorldMapTitle(header) {
        header.append('p').text(sanitizeHtml(this.settings.locations.title));
    }

    private getUSAStateCodes() {
        return [
            {
                state: 'Arizona',
                code: 'AZ'
            }, {
                state: 'Florida',
                code: 'FL'
            }, {
                state: 'Georgia',
                code: 'GA'
            }, {
                state: 'Illinois',
                code: 'IL'
            }, {
                state: 'Los Angeles',
                code: 'LA'
            }, {
                state: 'Maryland',
                code: 'MD'
            }, {
                state: 'Maine',
                code: 'ME'
            }, {
                state: 'Missouri',
                code: 'MO'
            }, {
                state: 'North Carolina',
                code: 'NC'
            }, {
                state: 'Nebraska',
                code: 'NE'
            }, {
                state: 'Nevada',
                code: 'NV'
            }, {
                state: 'New Jersey',
                code: 'NJ'
            }, {
                state: 'New York',
                code: 'NY'
            }, {
                state: 'Ohio',
                code: 'OH'
            }, {
                state: 'Oklahoma',
                code: 'OK'
            }, {
                state: 'Pennsylvania',
                code: 'PA'
            }, {
                state: 'South Carolina',
                code: 'SC'
            }, {
                state: 'South Dakota',
                code: 'SD'
            }, {
                state: 'Tennessee',
                code: 'TN'
            }, {
                state: 'Texas',
                code: 'TX'
            }, {
                state: 'Utah',
                code: 'UT'
            }, {
                state: 'Wisconsin',
                code: 'WI'
            }, {
                state: 'Virginia',
                code: 'VA'
            }, {
                state: 'Washington',
                code: 'WA'
            }, {
                state: 'California',
                code: 'CA'
            }, {
                state: 'Connecticut',
                code: 'CT'
            }, {
                state: 'Alaska',
                code: 'AK'
            }, {
                state: 'Arkansas',
                code: 'AR'
            }, {
                state: 'Alabama',
                code: 'AL'
            }
        ]
    }

    private getCountryCodes() {
        return [
            {
                country: 'France',
                code: 'FRA'
            }, {
                country: 'Germany',
                code: 'DEU'
            }, {
                country: 'Italy',
                code: 'ITA'
            }, {
                country: 'Spain',
                code: 'ESP'
            }, {
                country: 'United Kingdom',
                code: 'GBR'
            }, {
                country: 'USA',
                code: 'USA'
            }, {
                country: 'India',
                code: 'IND'
            }, {
                country: 'Japan',
                code: 'JPN'
            }, {
                country: 'Canada',
                code: 'CAN'
            }, {
                country: 'Russia',
                code: 'RUS'
            }, {
                country: 'Australia',
                code: 'AUS'
            }, {
                country: 'Brazil',
                code: 'BRA'
            }, {
                country: 'Argentina',
                code: 'ARG'
            }, {
                country: 'Mexico',
                code: 'MEX'
            }, {
                country: 'China',
                code: 'CHN'
            }, {
                country: 'Belgium',
                code: 'BEL'
            }, {
                country: 'Denmark',
                code: 'DNK'
            }, {
                country: 'Sweden',
                code: 'SWE'
            }, {
                country: 'Czechia',
                code: 'CZE'
            }, {
                country: 'Singapore',
                code: 'SGP'
            }, {
                country: 'South Korea',
                code: 'KOR'
            }, {
                country: 'Thailand',
                code: 'THA'
            }, {
                country: 'Turkey',
                code: 'TUR'
            }, {
                country: 'Saudi Arabia',
                code: 'SAU'
            }, {
                country: 'Egypt',
                code: 'EGY'
            }, {
                country: 'UAE',
                code: 'ARE'
            }
        ];
    }

    private getDefaultFills() {
        return {
            defaultFill: '#C9C9C9'
        };
    }

    private getDistinctYears(mapData) {
        return mapData.map(v => v.Launch).filter((v, i, list) => list.indexOf(v) === i).sort();;
    }

    private getYearColorData(mapData, distinctYears) {
        return distinctYears.map((v, i) => {
            let yearColor = mapData.find(y => y.Launch === v);
            return {
                Year: v,
                Color: yearColor.Color
            }
        });
    }

    private applyFills(fills) {
        this.yearColorData.forEach((v, i) => {
            if (v.Color) {
                fills[v.Color.toString()] = v.Color;
            }
        });
    }

    private getDatamapColorData(mapData, countryCodes) {
        let colorData = {};
        mapData.forEach((v, i) => {
            let countryCode = countryCodes.find((c, i) => c.country.toLowerCase() === v.Country.toLowerCase());
            let yearColor = this.yearColorData.find((y, i) => y.Year === v.Launch);
            if (countryCode && countryCode.code && yearColor && yearColor.Color) {
                colorData[countryCode.code] = { fillKey: yearColor.Color, selectionId: v.selectionId };
            }
        });
        return colorData;
    }

    private getDatamapColorDataFromUSA(mapData, stateCodes) {
        let colorData = {};
        mapData.forEach((v, i) => {
            let stateCode = stateCodes.find((c, i) => c.state.toLowerCase() === v.State.toLowerCase());
            let yearColor = this.yearColorData.find((y, i) => y.Year === v.Launch);
            if (stateCode && stateCode.code && yearColor && yearColor.Color) {
                colorData[stateCode.code] = { fillKey: yearColor.Color, selectionId: v.selectionId };
            }
        });
        return colorData;
    }

    private getSelectionId(mapData: any, country: string) {
        let selectionId;
        mapData.forEach((v) => {
            if (v.Country === country) {
                selectionId = v.selectionId;
            };
        })
        return selectionId;
    }

    private renderWorldMap(mapData, countryCodes, container) {
        let self = this;

        let mapHeight = 0;

        let fills = this.getDefaultFills();

        let distinctYears = this.getDistinctYears(mapData)

        this.yearColorData = this.getYearColorData(mapData, distinctYears);

        this.applyFills(fills);

        let data = this.getDatamapColorData(mapData, countryCodes);

        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
            mapHeight = 450;
        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
            mapHeight = 550;
        }

        let map = new Datamap({
            element: container.node(),
            projection: 'mercator',
            fills: fills,
            height: mapHeight, 
            data: data,
            done: (datamap) => {
                datamap.svg.data(mapData).selectAll('.datamaps-subunit').on('click', (geography) => {
                    let country = countryCodes.find((v, i) => v.code.toLowerCase() === geography.id.toLowerCase());
                    if (country && country.country) {
                        let doc = mapData.find((v, i) => v.Country === country.country);
                        if (doc && doc.DocumentLink) {
                            self.host.launchUrl(doc.DocumentLink);
                        }
                    }
                });
            }
        });
    }

    private renderWorldMapLegend(mainContent, options) {
        let legendContainer = mainContent.append('div')
            .attr('class', 'world-legend-container')
            .attr('style', 'height:' + (options.viewport.height - 60.13) + 'px;');

        let reversedYearColorData = this.yearColorData.map((v, i) => {
            let reversedYear;
            if (v.Year) {
                let splits = v.Year.split(' ');
                if (splits && splits.length === 2) {
                    reversedYear = splits[1] + ' ' + splits[0];
                }
                else {
                    reversedYear = splits[0];
                }
                return { Year: reversedYear, Color: v.Color };
            }
        });

        reversedYearColorData.sort((a, b) => {
            if (a.Year < b.Year)
                return -1;
            if (a.Year > b.Year)
                return 1;
            return 0;
        });

        let legend = legendContainer.selectAll('.legend')
            .data(reversedYearColorData)
            .enter()
            .append('div')
            .attr('class', 'legend');

        legend.append('div')
            .attr('class', 'color')
            .style('background-color', (d, i) => {
                return d.Color ? d.Color.toLowerCase() : '';
            });

        legend.append('div')
            .attr('class', 'year')
            .text((d, i) => {
                let splits = d.Year && d.Year.split(' ');
                if (splits && splits.length === 2) {
                    return splits[1] + ' ' + splits[0];
                }
                else {
                    return splits[0];
                }
            });
    }

    private getCurrentRegion(mapData) {
        let currentRegion = '';

        if (this.settings.locations.viewRegionalMap && this.settings.locations.defaultRegion === "USA") {
            currentRegion = this.settings.locations.defaultRegion;
        } else {
            let distinctRegions = mapData.map(d => d.Region).filter((v, i, list) => list.indexOf(v) === i);

            if (distinctRegions && distinctRegions.length === 1) {
                currentRegion = distinctRegions[0].toString();
            }
            else {
                currentRegion = this.settings.locations.defaultRegion;
            }
        }

        return currentRegion;
    }

    private renderRegionalMapTitle(header, currentRegion) {
        header.append('p').text(sanitizeHtml(this.settings.locations.title + ' - ' + currentRegion));
    }

    private createRegionMapWrapElement(containerWrap, options) {
        let regionMapWrap;
        if (this.settings.locations.viewHighlights) {
            regionMapWrap = containerWrap.append('div')
                .attr('class', 'region-map-wrap')
                .style('position', 'relative')
                .style('width', ((options.viewport.width * 70 / 100)) + 'px');
        } else {
            regionMapWrap = containerWrap.append('div')
                .attr('class', 'region-map-wrap')
                .style('position', 'relative')
                .style('width', (options.viewport.width) + 'px');
        }
        return regionMapWrap;
    }

    private createRegionMapElement(regionMapWrap, options) {
        let regionMap;
        let regionWidth = options.viewport.width;
        let regionHeight = options.viewport.height - 105;

        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
            regionWidth = (this.settings.locations.defaultRegion === "USA") ? (options.viewport.width - 300) : (options.viewport.width - 595);
            regionHeight = options.viewport.height - 595;
        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
            regionWidth = (this.settings.locations.defaultRegion === "USA") ? (options.viewport.width - 150) : regionWidth;
            regionHeight = options.viewport.height - 195;
        }

        if (this.settings.locations.viewHighlights) {
            let regionMapWidth = (this.settings.locations.headerImage && this.settings.locations.footerImage) ? regionWidth : (regionWidth - 395);
            regionMapWidth = (this.settings.locations.defaultRegion === "USA") ? regionWidth : regionMapWidth;
            regionMap = regionMapWrap.append('div')
                .attr('class', 'region-map')
                .style('height', (regionHeight) + 'px')
                .style('width', (regionMapWidth) + 'px');
        } else {
            regionMap = regionMapWrap.append('div')
                .attr('class', 'region-map')
                .style('height', (regionHeight) + 'px')
                .style('width', (regionWidth) + 'px');
        }
        return regionMap;
    }

    private getUSAStateMap(mapData, regionMap, fills, self) {

        let USAStateHeight = 0;

        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
            USAStateHeight = 350;
        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
            USAStateHeight = 450;
        }

        let stateCodes = this.getUSAStateCodes();

            let data = this.getDatamapColorDataFromUSA(mapData, stateCodes);

            let map = new Datamap({
                element: regionMap.node(),
                scope: 'usa',
                fills: fills,
                data: data,
                height: USAStateHeight,
                done: (datamap) => {
                    datamap.svg.attr("style", "margin-top:50;height:590;overflow: inherit;");
                    datamap.svg.data(mapData).selectAll('.datamaps-subunit').on('click', (geography) => {
                        let state = stateCodes.find((v, i) => v.code.toLowerCase() === geography.id.toLowerCase());
                        if (state && state.state) {
                            let document = mapData.find((v, i) => v.State === state.state);
                            if (document && document.DocumentLink) {
                                self.host.launchUrl(document.DocumentLink);
                            }
                        }
                    });
                }
            });
    }

    private renderRegionalMap(mapData, countryCodes, regionMap, currentRegion) {
        let self = this;
        let fills = this.getDefaultFills();
        let distinctYears = mapData.map(v => v.Launch).filter((v, i, list) => list.indexOf(v) === i).sort((a: any, b: any) => a - b);
        this.yearColorData = this.getYearColorData(mapData, distinctYears);
        this.applyFills(fills);
        let regionHeight = 0;
        if (this.settings.locations.headerImage && this.settings.locations.footerImage) { regionHeight = 390 } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) { regionHeight = 490 };

        if (currentRegion !=  "USA") {

            let data = this.getDatamapColorData(mapData, countryCodes);

            let map = new Datamap({
                element: regionMap.node(),
                scope: 'world',
                setProjection: (element) => {
                    let coordinationX,coordinationY,zoom;
                    if (currentRegion === 'Europe') {
                        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
                            coordinationX = 15.2551;coordinationY = 75;zoom = 225;
                        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
                            coordinationX = 15.2551;coordinationY = 58;zoom = 325;
                        } else {
                            coordinationX = 15.2551;coordinationY = 58;zoom = 425;
                        }
                    } else if (currentRegion === 'Asia') {
                        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
                            coordinationX = 100;coordinationY = 55;zoom = 215;
                        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
                            coordinationX = 160;coordinationY = 30;zoom = 250;
                            if (this.settings.locations.viewHighlights) { coordinationX = 130;coordinationY = 30;zoom = 220; }
                        } else {
                            coordinationX = 125;coordinationY = 30;zoom = 325;
                        }
                    }
                    else if (currentRegion === 'Lat-Am') {

                        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
                            coordinationX = -50;coordinationY = 10;zoom = 210;
                        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
                            coordinationX = -75;coordinationY = -25;zoom = 300;
                        } else {
                            coordinationX = -60;coordinationY = -25;zoom = 350;
                        }
                    } else if (currentRegion === 'NA') {
                        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
                            coordinationX = -110;coordinationY = 82;zoom = 120;
                        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
                            coordinationX = -110;coordinationY = 67;zoom = 140;
                        } else {
                            coordinationX = -110;coordinationY = 67;zoom = 190;
                        }
                    }
                    else if (currentRegion === 'AfME') {
                        if (this.settings.locations.headerImage && this.settings.locations.footerImage) {
                            coordinationX = 30;coordinationY = 30;zoom = 270;
                        } else if (this.settings.locations.headerImage || this.settings.locations.footerImage) {
                            coordinationX = 30;coordinationY = 2;zoom = 340;
                        } else {
                            coordinationX = 30;coordinationY = 4;zoom = 430;
                        }
                    }
                    if (zoom) {
                        let projection = d3.geoMercator()
                            .center([coordinationX, coordinationY])
                            .scale(zoom)
                            .translate([element.offsetWidth / 2, element.offsetHeight / 2]);
                        let path = d3.geoPath()
                            .projection(projection);
                        return { path: path, projection: projection };
                    }
                },
                fills: fills,
                height: regionHeight,
                data: data,
                done: (datamap) => {
                    datamap.svg.data(mapData).selectAll('.datamaps-subunit').on('click', (geography) => {
                        let country = countryCodes.find((v, i) => v.code.toLowerCase() === geography.id.toLowerCase());
                        if (country && country.country) {
                            let document = mapData.find((v, i) => v.Country === country.country);
                            if (document && document.DocumentLink) {
                                self.host.launchUrl(document.DocumentLink);
                            }
                        }
                    });
                }
            });
        } else if (currentRegion === "USA") {

            this.getUSAStateMap(mapData,regionMap, fills, self);

        }

    }

    private createHighlightsContainerElement(containerWrap, mapData: GlobalFacilityLocation[]) {
        if (this.settings.locations.viewHighlights) {
            let highlights = containerWrap.append('div')
                .attr('class', 'highlights');

            highlights.append('div')
                .attr('class', 'highlights-header')
                .text('Highlights');

            // here we are using html method because the column or property Highlights has value as HTML content (Rich text)
            let [map] = mapData;
            if (map.Highlights) {
                highlights.append('div')
                .attr('class', 'highlights-content')
                .html(sanitizeHtml(map.Highlights) ? sanitizeHtml(map.Highlights.toString()) : '');
            }
        }
    }

    private renderRegionalMapLegend(mainContent, currentRegion) {
        let legendContainer;
        if (currentRegion !== "USA") {
            legendContainer = mainContent.append('div')
            .attr('class', 'regional-legend-container')
        } else if (currentRegion === "USA") {
            legendContainer = mainContent.append('div')
            .attr('class', 'regional-legend-container').attr("style", "bottom:5px;left:5px;flex-direction: column;justify-content: flex-end;")
        }

        let legend = legendContainer.selectAll('.legend')
            .data(this.yearColorData)
            .enter()
            .append('div')
            .attr('class', 'legend');

        legend.append('div')
            .attr('class', 'color')
            .style('background-color', (d, i) => {
                return d.Color ? d.Color.toLowerCase() : '';
            });

        legend.append('div')
            .attr('class', 'year')
            .text((d, i) => {
                return d.Year;
            });
    }

    // converter to table data
    public static CONVERTER(dataView: DataView, host: IVisualHost): GlobalFacilityLocation[] {
        let resultData: GlobalFacilityLocation[] = [];
        let tableView = dataView.table;
        let _rows = tableView.rows;
        let _columns = tableView.columns;
        let _companyIndex = -1, _regionIndex = -1, _stateIndex = -1, _countryIndex = -1, _docIndex = -1,
            _launchIndex = -1, _colorIndex = -1, _highlightsIndex = -1, _headerImageIndex = -1, _footerImageIndex = -1;
        for (let ti = 0; ti < _columns.length; ti++) {
            if (_columns[ti].roles.hasOwnProperty("Company")) {
                _companyIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Region")) {
                _regionIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("State")) {
                _stateIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Country")) {
                _countryIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("DocumentLink")) {
                _docIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Launch")) {
                _launchIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Color")) {
                _colorIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Highlights")) {
                _highlightsIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("HeaderImage")) {
                _headerImageIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("FooterImage")) {
                _footerImageIndex = ti;
            }
        }
        for (let i = 0; i < _rows.length; i++) {
            let row = _rows[i];
            let dp = {
                Company: row[_companyIndex] ? row[_companyIndex].toString() : null,
                Region: row[_regionIndex] ? row[_regionIndex].toString() : null,
                State: row[_stateIndex] ? row[_stateIndex].toString() : null,
                Country: row[_countryIndex] ? row[_countryIndex].toString() : null,
                DocumentLink: row[_docIndex] ? row[_docIndex].toString() : null,
                Launch: row[_launchIndex] ? row[_launchIndex].toString() : null,
                Color: row[_colorIndex] ? row[_colorIndex].toString() : null,
                Highlights: row[_highlightsIndex] ? row[_highlightsIndex].toString() : null,
                HeaderImage: row[_headerImageIndex] ? row[_headerImageIndex].toString() : null,
                FooterImage: row[_footerImageIndex] ? row[_footerImageIndex].toString() : null,
                selectionId:host.createSelectionIdBuilder().withTable(tableView, i).createSelectionId()
            };
            resultData.push(dp);
        }
        return resultData;
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return <VisualSettings>VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }
}