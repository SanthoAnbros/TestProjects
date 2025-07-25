import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
import { map, catchError } from 'rxjs/operators';
import { Employee, EmployeeCreateRequest, EmployeeUpdateRequest } from '../models/employee.model';

@Injectable({
  providedIn: 'root'
})
export class SharePointService {
  private siteUrl = 'https://yourtenant.sharepoint.com/sites/yoursite'; // Replace with your site URL
  private listName = 'employee';
  private apiUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;

  constructor(private http: HttpClient) { }

  private getHeaders(): HttpHeaders {
    return new HttpHeaders({
      'Accept': 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': this.getRequestDigest()
    });
  }

  private getRequestDigest(): string {
    // In a real application, you would get this from SharePoint context
    // For demo purposes, this is a placeholder
    const element = document.querySelector<HTMLInputElement>('#__REQUESTDIGEST');
    return element ? element.value : '';
  }

  // Get all employees
  getAllEmployees(): Observable<Employee[]> {
    const url = `${this.apiUrl}?$select=Id,Title,FirstName,LastName,Email,Department,Position,Phone,HireDate,Salary,Status,Created,Modified&$orderby=Created desc`;
    
    return this.http.get(url, { headers: this.getHeaders() }).pipe(
      map((response: any) => {
        return response.d.results.map((item: any) => ({
          Id: item.Id,
          Title: item.Title,
          FirstName: item.FirstName,
          LastName: item.LastName,
          Email: item.Email,
          Department: item.Department,
          Position: item.Position,
          Phone: item.Phone,
          HireDate: new Date(item.HireDate),
          Salary: item.Salary,
          Status: item.Status,
          Created: new Date(item.Created),
          Modified: new Date(item.Modified)
        }));
      }),
      catchError(error => {
        console.error('Error fetching employees:', error);
        return throwError(error);
      })
    );
  }

  // Get employee by ID
  getEmployeeById(id: number): Observable<Employee> {
    const url = `${this.apiUrl}(${id})?$select=Id,Title,FirstName,LastName,Email,Department,Position,Phone,HireDate,Salary,Status,Created,Modified`;
    
    return this.http.get(url, { headers: this.getHeaders() }).pipe(
      map((response: any) => ({
        Id: response.d.Id,
        Title: response.d.Title,
        FirstName: response.d.FirstName,
        LastName: response.d.LastName,
        Email: response.d.Email,
        Department: response.d.Department,
        Position: response.d.Position,
        Phone: response.d.Phone,
        HireDate: new Date(response.d.HireDate),
        Salary: response.d.Salary,
        Status: response.d.Status,
        Created: new Date(response.d.Created),
        Modified: new Date(response.d.Modified)
      })),
      catchError(error => {
        console.error('Error fetching employee:', error);
        return throwError(error);
      })
    );
  }

  // Create new employee
  createEmployee(employee: EmployeeCreateRequest): Observable<Employee> {
    const headers = this.getHeaders();
    
    const body = {
      __metadata: { type: 'SP.Data.EmployeeListItem' }, // Adjust based on your list internal name
      Title: employee.Title,
      FirstName: employee.FirstName,
      LastName: employee.LastName,
      Email: employee.Email,
      Department: employee.Department,
      Position: employee.Position,
      Phone: employee.Phone || '',
      HireDate: employee.HireDate.toISOString(),
      Salary: employee.Salary || 0,
      Status: employee.Status
    };

    return this.http.post(this.apiUrl, body, { headers }).pipe(
      map((response: any) => ({
        Id: response.d.Id,
        Title: response.d.Title,
        FirstName: response.d.FirstName,
        LastName: response.d.LastName,
        Email: response.d.Email,
        Department: response.d.Department,
        Position: response.d.Position,
        Phone: response.d.Phone,
        HireDate: new Date(response.d.HireDate),
        Salary: response.d.Salary,
        Status: response.d.Status,
        Created: new Date(response.d.Created),
        Modified: new Date(response.d.Modified)
      })),
      catchError(error => {
        console.error('Error creating employee:', error);
        return throwError(error);
      })
    );
  }

  // Update employee
  updateEmployee(employee: EmployeeUpdateRequest): Observable<void> {
    const url = `${this.apiUrl}(${employee.Id})`;
    const headers = this.getHeaders().set('X-HTTP-Method', 'MERGE').set('IF-MATCH', '*');
    
    const body = {
      __metadata: { type: 'SP.Data.EmployeeListItem' },
      Title: employee.Title,
      FirstName: employee.FirstName,
      LastName: employee.LastName,
      Email: employee.Email,
      Department: employee.Department,
      Position: employee.Position,
      Phone: employee.Phone || '',
      HireDate: employee.HireDate.toISOString(),
      Salary: employee.Salary || 0,
      Status: employee.Status
    };

    return this.http.post(url, body, { headers }).pipe(
      map(() => {}),
      catchError(error => {
        console.error('Error updating employee:', error);
        return throwError(error);
      })
    );
  }

  // Delete employee
  deleteEmployee(id: number): Observable<void> {
    const url = `${this.apiUrl}(${id})`;
    const headers = this.getHeaders().set('X-HTTP-Method', 'DELETE').set('IF-MATCH', '*');

    return this.http.post(url, {}, { headers }).pipe(
      map(() => {}),
      catchError(error => {
        console.error('Error deleting employee:', error);
        return throwError(error);
      })
    );
  }

  // Search employees
  searchEmployees(searchTerm: string): Observable<Employee[]> {
    const filter = `substringof('${searchTerm}',Title) or substringof('${searchTerm}',FirstName) or substringof('${searchTerm}',LastName) or substringof('${searchTerm}',Email) or substringof('${searchTerm}',Department) or substringof('${searchTerm}',Position)`;
    const url = `${this.apiUrl}?$select=Id,Title,FirstName,LastName,Email,Department,Position,Phone,HireDate,Salary,Status,Created,Modified&$filter=${filter}&$orderby=Created desc`;
    
    return this.http.get(url, { headers: this.getHeaders() }).pipe(
      map((response: any) => {
        return response.d.results.map((item: any) => ({
          Id: item.Id,
          Title: item.Title,
          FirstName: item.FirstName,
          LastName: item.LastName,
          Email: item.Email,
          Department: item.Department,
          Position: item.Position,
          Phone: item.Phone,
          HireDate: new Date(item.HireDate),
          Salary: item.Salary,
          Status: item.Status,
          Created: new Date(item.Created),
          Modified: new Date(item.Modified)
        }));
      }),
      catchError(error => {
        console.error('Error searching employees:', error);
        return throwError(error);
      })
    );
  }
}