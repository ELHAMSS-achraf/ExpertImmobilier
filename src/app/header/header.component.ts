import { UserService } from './../ServicesFolder/user.service';
import { Component, Input, OnInit, Output, EventEmitter } from '@angular/core';
import { Router } from '@angular/router';
import { IDanger } from '../Models/idanger';
import { Logged } from '../Models/logged';





@Component({
  selector: 'app-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.css']
})


export class HeaderComponent implements OnInit {
   currentPage :string = "";
   authUsers : Logged[] = [];
   loggedIn : boolean = false;
   @Output() evtLogOut = new EventEmitter();
   @Input() newPoint : IDanger = {
    id : 0,
    user : {
      id : 0,
      name : "",
      isLogged : false,
    },
    type : "",
    date :new Date (0,0,0),
    commentaire :"",
    longitude :50,
    latitude:100
  }
  point = this.newPoint;

  constructor(private router: Router, private _userService : UserService) { 
    this.changeHeader();
    
        
  }
 
  
  ngOnInit(): void {
    this._userService.getLogged()
    .subscribe(data => this.authUsers=data);
  }


  
  logOut(){
    this.evtLogOut.emit();
  }

  changeHeader(){
      switch (this.router.url) {
        case "/Register":
            this.currentPage = "register";
          break;
        case "/Login":
            this.currentPage = "login";
            break;
        case "/Map":
          this.currentPage = "map";
        break;        
        
      }
  }

  log(){
    this.point =this.newPoint;
  }




}
