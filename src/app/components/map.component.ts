import { Component, OnInit, ChangeDetectionStrategy, Input, ElementRef } from '@angular/core';
import Map from 'ol/Map';

@Component({
  selector: 'app-map',
  template: '<main><div class="left-bar"></div><div class="map-container"><app-map [map]="map"></app-map></div></main>',
  styles: [':host { width: 100%; height: 100%; display: block; }',
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class MapComponent implements OnInit {
  @Input() map?: Map;
  constructor(private elementRef: ElementRef) {
  }
  ngOnInit() {
    this.map?.setTarget(this.elementRef.nativeElement);
  }
}