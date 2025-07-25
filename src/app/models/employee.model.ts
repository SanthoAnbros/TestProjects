export interface Employee {
  Id?: number;
  Title: string;
  FirstName: string;
  LastName: string;
  Email: string;
  Department: string;
  Position: string;
  Phone?: string;
  HireDate: Date;
  Salary?: number;
  Status: 'Active' | 'Inactive';
  Created?: Date;
  Modified?: Date;
}

export interface EmployeeCreateRequest {
  Title: string;
  FirstName: string;
  LastName: string;
  Email: string;
  Department: string;
  Position: string;
  Phone?: string;
  HireDate: Date;
  Salary?: number;
  Status: 'Active' | 'Inactive';
}

export interface EmployeeUpdateRequest extends EmployeeCreateRequest {
  Id: number;
}