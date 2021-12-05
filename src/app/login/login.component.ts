import { Component, Input, OnInit } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { Logged } from '../Models/logged';
import { User } from '../Models/user';
import { UserService } from '../ServicesFolder/user.service';

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.css']
})
export class LoginComponent implements OnInit {

  /*   initiate authentificated statut  with false*/  
  authentificated=true;
  
  public users: User[] = [];
  public logged: Logged ={
    id :0,
    name : "",
    isLogged :false
  };

  userForm: FormGroup;
  userModel : User={
    id:0,
    name:'',
    email:'',
    password:''
  };

  constructor(private _userService:UserService,private fb:FormBuilder,private router: Router) {

    this.userForm=this.fb.group(
      {
        name:[''],
        password:['']
      }
    )
   }

  ngOnInit(): void {
    this._userService.getUsers().subscribe(data=> this.users=data);
  }


  logIn(){
    this.userModel=this.userForm.value;
    console.log(this.userForm.value);

    
    this.users.forEach( user =>{
      if(user.name==this.userModel.name && user.password==this.userModel.password){

        this.userForm.reset();
        this.logged.id = user.id;
        this.logged.name = user.name;
        this.logged.isLogged = true;
        this.router.navigateByUrl('/report');
      }
    }); 
    if (!this.logged.isLogged){
      alert("Error");
      this.userForm.reset();
    }
}



}
