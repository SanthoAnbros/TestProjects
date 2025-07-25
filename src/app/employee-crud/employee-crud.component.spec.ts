import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { EmployeeCrudComponent } from './employee-crud.component';

describe('EmployeeCrudComponent', () => {
  let component: EmployeeCrudComponent;
  let fixture: ComponentFixture<EmployeeCrudComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ EmployeeCrudComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(EmployeeCrudComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
