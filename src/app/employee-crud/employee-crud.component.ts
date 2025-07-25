import { Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { SharePointService } from '../services/sharepoint.service';
import { Employee, EmployeeCreateRequest, EmployeeUpdateRequest } from '../models/employee.model';

@Component({
  selector: 'app-employee-crud',
  templateUrl: './employee-crud.component.html',
  styleUrls: ['./employee-crud.component.css']
})
export class EmployeeCrudComponent implements OnInit {
  employees: Employee[] = [];
  filteredEmployees: Employee[] = [];
  employeeForm: FormGroup;
  isEditMode = false;
  isLoading = false;
  showForm = false;
  searchTerm = '';
  selectedEmployeeId: number | null = null;
  
  departments = ['HR', 'IT', 'Finance', 'Marketing', 'Operations', 'Sales'];
  statuses = ['Active', 'Inactive'];

  constructor(
    private fb: FormBuilder,
    private sharePointService: SharePointService
  ) {
    this.employeeForm = this.createForm();
  }

  ngOnInit(): void {
    this.loadEmployees();
  }

  createForm(): FormGroup {
    return this.fb.group({
      title: ['', [Validators.required, Validators.maxLength(100)]],
      firstName: ['', [Validators.required, Validators.maxLength(50)]],
      lastName: ['', [Validators.required, Validators.maxLength(50)]],
      email: ['', [Validators.required, Validators.email]],
      department: ['', Validators.required],
      position: ['', [Validators.required, Validators.maxLength(100)]],
      phone: ['', [Validators.pattern(/^[\+]?[1-9][\d]{0,15}$/)]],
      hireDate: ['', Validators.required],
      salary: ['', [Validators.min(0)]],
      status: ['Active', Validators.required]
    });
  }

  loadEmployees(): void {
    this.isLoading = true;
    this.sharePointService.getAllEmployees().subscribe({
      next: (employees) => {
        this.employees = employees;
        this.filteredEmployees = employees;
        this.isLoading = false;
      },
      error: (error) => {
        console.error('Error loading employees:', error);
        this.isLoading = false;
        // In a real application, show proper error message to user
        alert('Error loading employees. Please check console for details.');
      }
    });
  }

  onSearch(): void {
    if (this.searchTerm.trim() === '') {
      this.filteredEmployees = this.employees;
    } else {
      this.sharePointService.searchEmployees(this.searchTerm).subscribe({
        next: (employees) => {
          this.filteredEmployees = employees;
        },
        error: (error) => {
          console.error('Error searching employees:', error);
          // Fallback to local filtering
          this.filteredEmployees = this.employees.filter(emp =>
            emp.Title.toLowerCase().includes(this.searchTerm.toLowerCase()) ||
            emp.FirstName.toLowerCase().includes(this.searchTerm.toLowerCase()) ||
            emp.LastName.toLowerCase().includes(this.searchTerm.toLowerCase()) ||
            emp.Email.toLowerCase().includes(this.searchTerm.toLowerCase()) ||
            emp.Department.toLowerCase().includes(this.searchTerm.toLowerCase()) ||
            emp.Position.toLowerCase().includes(this.searchTerm.toLowerCase())
          );
        }
      });
    }
  }

  showAddForm(): void {
    this.isEditMode = false;
    this.showForm = true;
    this.employeeForm.reset();
    this.employeeForm.patchValue({ status: 'Active' });
  }

  editEmployee(employee: Employee): void {
    this.isEditMode = true;
    this.showForm = true;
    this.selectedEmployeeId = employee.Id!;
    
    this.employeeForm.patchValue({
      title: employee.Title,
      firstName: employee.FirstName,
      lastName: employee.LastName,
      email: employee.Email,
      department: employee.Department,
      position: employee.Position,
      phone: employee.Phone,
      hireDate: this.formatDateForInput(employee.HireDate),
      salary: employee.Salary,
      status: employee.Status
    });
  }

  deleteEmployee(id: number): void {
    if (confirm('Are you sure you want to delete this employee?')) {
      this.isLoading = true;
      this.sharePointService.deleteEmployee(id).subscribe({
        next: () => {
          this.loadEmployees();
          alert('Employee deleted successfully!');
        },
        error: (error) => {
          console.error('Error deleting employee:', error);
          this.isLoading = false;
          alert('Error deleting employee. Please try again.');
        }
      });
    }
  }

  onSubmit(): void {
    if (this.employeeForm.valid) {
      this.isLoading = true;
      const formValue = this.employeeForm.value;
      
      if (this.isEditMode && this.selectedEmployeeId) {
        const updateRequest: EmployeeUpdateRequest = {
          Id: this.selectedEmployeeId,
          Title: formValue.title,
          FirstName: formValue.firstName,
          LastName: formValue.lastName,
          Email: formValue.email,
          Department: formValue.department,
          Position: formValue.position,
          Phone: formValue.phone,
          HireDate: new Date(formValue.hireDate),
          Salary: formValue.salary,
          Status: formValue.status
        };

        this.sharePointService.updateEmployee(updateRequest).subscribe({
          next: () => {
            this.loadEmployees();
            this.cancelForm();
            alert('Employee updated successfully!');
          },
          error: (error) => {
            console.error('Error updating employee:', error);
            this.isLoading = false;
            alert('Error updating employee. Please try again.');
          }
        });
      } else {
        const createRequest: EmployeeCreateRequest = {
          Title: formValue.title,
          FirstName: formValue.firstName,
          LastName: formValue.lastName,
          Email: formValue.email,
          Department: formValue.department,
          Position: formValue.position,
          Phone: formValue.phone,
          HireDate: new Date(formValue.hireDate),
          Salary: formValue.salary,
          Status: formValue.status
        };

        this.sharePointService.createEmployee(createRequest).subscribe({
          next: () => {
            this.loadEmployees();
            this.cancelForm();
            alert('Employee created successfully!');
          },
          error: (error) => {
            console.error('Error creating employee:', error);
            this.isLoading = false;
            alert('Error creating employee. Please try again.');
          }
        });
      }
    } else {
      this.markFormGroupTouched();
    }
  }

  cancelForm(): void {
    this.showForm = false;
    this.isEditMode = false;
    this.selectedEmployeeId = null;
    this.employeeForm.reset();
    this.isLoading = false;
  }

  private markFormGroupTouched(): void {
    Object.keys(this.employeeForm.controls).forEach(key => {
      const control = this.employeeForm.get(key);
      if (control) {
        control.markAsTouched();
      }
    });
  }

  private formatDateForInput(date: Date): string {
    const d = new Date(date);
    const year = d.getFullYear();
    const month = (d.getMonth() + 1).toString().padStart(2, '0');
    const day = d.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  isFieldInvalid(fieldName: string): boolean {
    const field = this.employeeForm.get(fieldName);
    return !!(field && field.invalid && (field.dirty || field.touched));
  }

  getFieldError(fieldName: string): string {
    const field = this.employeeForm.get(fieldName);
    if (field && field.errors) {
      if (field.errors['required']) return `${fieldName} is required`;
      if (field.errors['email']) return 'Invalid email format';
      if (field.errors['maxlength']) return `${fieldName} is too long`;
      if (field.errors['min']) return `${fieldName} must be greater than 0`;
      if (field.errors['pattern']) return `Invalid ${fieldName} format`;
    }
    return '';
  }
}
