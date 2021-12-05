import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { IDanger } from '../Models/idanger';

@Injectable({
  providedIn: 'root'
})
export class DangerService {
  private _url:string="http://localhost:3000/dangers";
 


  constructor(private http : HttpClient) { }
  addDanger(danger :IDanger) : Observable<IDanger>{
    return this.http.post<IDanger>(this._url,danger);

  }


  getDangers():Observable<IDanger[]>{
    return this.http.get<IDanger[]>(this._url);
  }

}
