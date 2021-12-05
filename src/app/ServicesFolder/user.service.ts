import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { User } from '../Models/user';
import { Logged } from '../Models/logged';


@Injectable({
  providedIn: 'root'
})
export class UserService {
  private _url:string="http://localhost:3000/users";

  constructor(private http : HttpClient) { }
  
  addUser(user :User) : Observable<User>{
    return this.http.post<User>(this._url,user);

  }

  getUsers():Observable<User[]>{
    return this.http.get<User[]>(this._url);
  }



 



 




}
