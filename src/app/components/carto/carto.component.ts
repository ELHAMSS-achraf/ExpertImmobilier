import { Component, OnInit, ElementRef } from '@angular/core';
import Map from "@arcgis/core/Map";
import MapView from '@arcgis/core/views/MapView';
import BasemapGallery from "@arcgis/core/widgets/BasemapGallery";
import Expand from "@arcgis/core/widgets/Expand";
import Fullscreen from "@arcgis/core/widgets/Fullscreen";
import Locate from "@arcgis/core/widgets/Locate";
import Search from "@arcgis/core/widgets/Search";

import Print from "@arcgis/core/widgets/Print";
@Component({
  selector: 'app-carto',
  templateUrl: './carto.component.html',
  styleUrls: ['./carto.component.css']
})
export class CartoComponent implements OnInit {


map :Map ;

  constructor() {
    this.map = new Map({
      basemap: "satellite"
    });
  }

    ngOnInit(): void {
      this.loadMap();
    
    }
    loadMap() {
      /* esriConfig.apiKey = "YOUR_API_KEY"; */
     
      const view = new MapView({
          container: 'mapDiv',
          map :this.map,
          zoom: 12,
         center: [-7.59, 33.58]
          
      });
  
      let basemapGallery = new BasemapGallery({
        view: view
      });
  
      let bgExpand = new Expand({
        view: view,
        content: basemapGallery
      });
  
      basemapGallery.watch("activeBasemap", function() {
        var mobileSize = view.heightBreakpoint === "xsmall" || view.widthBreakpoint === "xsmall";
  
        if (mobileSize) {
          bgExpand.collapse();
        }
      });
  
      view.ui.add(bgExpand, {
        position: "bottom-right"
      });
  
      let fullscreen = new Fullscreen({
        view: view
      });
  
      view.ui.add(fullscreen, "bottom-right");
  
      let locateWidget = new Locate({
        view: view,   // Attaches the Locate button to the view
        /* graphic: new Graphic({
          symbol: { type: "simple-marker" }  // overwrites the default symbol used for the
          // graphic placed at the location of the user when found
        }) */
      });
      
      view.ui.add(locateWidget, "bottom-right");
  
      const searchWidget = new Search({
        view: view
      });
  
      // Adds the search widget below other elements in
      view.ui.add(searchWidget, {
        position: "top-right",
        index: 2
      });
  
      const print = new Print({
        view: view,
        // specify your own print service
        printServiceUrl:
           "https://utility.arcgisonline.com/arcgis/rest/services/Utilities/PrintingTools/GPServer/Export%20Web%20Map%20Task",
           allowedFormats: ["jpg", "png8", "png32"]
          });
      
      // Adds widget below other elements in the top left corner of the view
      view.ui.add(print, {
        position: "top-left"
      });
     
  
  
      
      
      
  
   
      
    }
    

  }