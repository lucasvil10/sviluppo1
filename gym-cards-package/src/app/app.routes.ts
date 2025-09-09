import { Routes } from '@angular/router';
import { AllenamentoComponent } from './allenamento/allenamento.component';
import { HomeComponent } from './home/home.component';

export const routes: Routes = [
  { path: 'allenamento', component: AllenamentoComponent },
  { path: 'home', component: HomeComponent },
  { path: '', redirectTo: '/home', pathMatch: 'full' }
];
