import { ComponentFixture, TestBed } from '@angular/core/testing';

import { GeneratorReportComponent } from './generator-report.component';

describe('GeneratorReportComponent', () => {
  let component: GeneratorReportComponent;
  let fixture: ComponentFixture<GeneratorReportComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ GeneratorReportComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(GeneratorReportComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
