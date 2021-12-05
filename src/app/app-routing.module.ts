import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { CartoComponent } from './components/carto/carto.component';
import { ContactComponent } from './components/contact/contact.component';
import { HomeComponent } from './components/home/home.component';
import { RegisterComponent } from './register/register.component';
import { LoginComponent } from './login/login.component';
import { GeneratorReportComponent } from './report/generator-report/generator-report.component';

const routes: Routes = [
  {path: "" , component:HomeComponent},
  {path: "carto" , component:CartoComponent},
  {path: "contact" , component:ContactComponent},
  {path : 'Login',component : LoginComponent,pathMatch: 'full'},
  {path : 'Register',component : RegisterComponent,pathMatch: 'full'},
  {path: "report" , component: GeneratorReportComponent},
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
