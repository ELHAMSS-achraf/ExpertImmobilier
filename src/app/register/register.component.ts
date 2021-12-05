import { Component, OnInit } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { User } from '../Models/user';
import { UserService } from '../ServicesFolder/user.service';

@Component({
  selector: 'app-register',
  templateUrl: './register.component.html',
  styleUrls: ['./register.component.css']
})
export class RegisterComponent implements OnInit {

  userForm!: FormGroup;
  userModel : User={
    id:0,
    name:'',
    email:'',
    password:''
  };

  constructor(private _userService:UserService,private fb:FormBuilder, private router: Router) {
    this.userForm=this.fb.group(
      {
        id:[''] ,
        name:['',[Validators.required,Validators.minLength(2),Validators.maxLength(10)]],
        email:['',[Validators.required, Validators.pattern("[a-zA-Z]*@gmail.com")]],
        password:['',[Validators.required,Validators.pattern("[a-zA-z0-9]{6,}")]]
      }
    )
   }
  

  ngOnInit(): void {}
  onSubmit(){
    if (this.userForm.valid) {
      this.userModel=this.userForm.value;   
      this._userService.addUser(this.userModel)//attention !!!!!!!!!!!!!!!!!
        .subscribe(
           data=> console.log('success data!',data)
           );
      
    } else {
      alert("please enter validate fields !");
    }
  
    

    
}

}
