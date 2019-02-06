/*
 * Copyright 2019 Coordinador Electrico Nacional
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cl.coordinador.peajes;

import java.text.DecimalFormat;

/**
 * Implementa las funcionalidades para manipular matrices
 *
 * @author
 */
public class Matriz {


   private float[][] A;

   private int m, n;

/* ------------------------
   Constructores
 * ------------------------ */

   /** Construct an m-by-n matrix of zeros.*/

   public Matriz (int mm, int nn) {
      m = mm;
      n = nn;
      A = new float[m][n];
   }

   /** Construct a matrix from a 2-D array.*/

   public Matriz (float[][] MatA) {
      m = MatA.length;
      n = MatA[0].length;
      for (int i = 0; i < m; i++) {
         if (MatA[i].length != n) {
            System.out.println(i);
            throw new IllegalArgumentException("All rows must have the same length.");
         }
      }
      A = MatA;
   }
   
   
/* ------------------------
   Metodos
 * ------------------------ */
   
   /** Linear algebraic matrix multiplication, A * B. Ignores zero elements
   @param B    another matrix
   @return     Matrix product, A * B
   @exception  IllegalArgumentException Matrix inner dimensions must agree.
   */
   public Matriz ProductoD (Matriz B) {
      if (B.m != n) {
         throw new IllegalArgumentException("Matrix inner dimensions must agree.");
      }
      Matriz X = new Matriz(m,B.n);
      float[][] C = X.getArray();
      float[] Bcolj = new float[n];
      for (int j = 0; j < B.n; j++) {
         for (int k = 0; k < n; k++) {
            Bcolj[k] = B.A[k][j];
         }
         for (int i = 0; i < m; i++) {
            float[] Arowi = A[i];
            float s = 0;
            for (int k = 0; k < n; k++) {
               if(Arowi[k]!=0 & Bcolj[k]!=0)
               		s += Arowi[k]*Bcolj[k];
            }
            C[i][j] = s;
         }
      }
      return X;
   }
   
   /** Access the internal two-dimensional array.
   @return     Pointer to the two-dimensional array of matrix elements.
   */
   public float[][] getArray () {
      return A;
   }
   
    /** Copy the internal two-dimensional array.
   @return     Two-dimensional array copy of matrix elements.
   */
   public float[][] getArrayCopy () {
      float[][] C = new float[m][n];
      for (int i = 0; i < m; i++) {
         for (int j = 0; j < n; j++) {
            C[i][j] = A[i][j];
         }
      }
      return C;
   }
   
   /** Get row dimension.
   @return     m, the number of rows.
   */
   public int numeroFil () {
      return m;
   }

   /** Get column dimension.
   @return     n, the number of columns.
   */
   public int numeroCol () {
      return n;
   }
   
   /** Get a submatrix.
   @param i0   Initial row index
   @param i1   Final row index
   @param j0   Initial column index
   @param j1   Final column index
   @return     A(i0:i1,j0:j1)
   @exception  ArrayIndexOutOfBoundsException Submatrix indices
   */
   public Matriz ObtenerSubMatriz (int i0, int i1, int j0, int j1) {
      Matriz X = new Matriz(i1-i0+1,j1-j0+1);
      float[][] B = X.getArray();
      try {
         for (int i = i0; i <= i1; i++) {
            for (int j = j0; j <= j1; j++) {
               B[i-i0][j-j0] = A[i][j];
            }
         }
      } catch(ArrayIndexOutOfBoundsException e) {
         throw new ArrayIndexOutOfBoundsException("Submatrix indices");
      }
      return X;
   }
   
   /** Set a submatrix.
   @param i0   Initial row index
   @param i1   Final row index
   @param j0   Initial column index
   @param j1   Final column index
   @param X    A(i0:i1,j0:j1)
   @exception  ArrayIndexOutOfBoundsException Submatrix indices
   */
   public void SetearSubMatriz (int i0, int i1, int j0, int j1, Matriz X) {
      try {
         for (int i = i0; i <= i1; i++) {
            for (int j = j0; j <= j1; j++) {
               A[i][j] = X.get(i-i0,j-j0);
            }
         }
      } catch(ArrayIndexOutOfBoundsException e) {
         throw new ArrayIndexOutOfBoundsException("Submatrix indices");
      }
   }
   
   /** Get a single element.
   @param i    Row index.
   @param j    Column index.
   @return     A(i,j)
   @exception  ArrayIndexOutOfBoundsException
   */
   public float get (int i, int j) {
      return A[i][j];
   }
   
   /** Set a single element.
   @param i    Row index.
   @param j    Column index.
   @param s    A(i,j).
   @exception  ArrayIndexOutOfBoundsException
   */
   public void set (int i, int j, float s) {
      A[i][j] = s;
   }
   
   /** C = A + B
   @param B    another matrix
   @return     A + B
   */
   public Matriz plus (Matriz B) {
      checkMatrixDimensions(B);
      Matriz X = new Matriz(m,n);
      float[][] C = X.getArray();
      for (int i = 0; i < m; i++) {
         for (int j = 0; j < n; j++) {
            if(A[i][j]!=0 | B.A[i][j]!=0)
               C[i][j] = A[i][j] + B.A[i][j];
         }
      }
      return X;
   }
   
   /** C = A - B
   @param B    another matrix
   @return     A - B
   */
   public Matriz minus (Matriz B) {
      checkMatrixDimensions(B);
      Matriz X = new Matriz(m,n);
      float[][] C = X.getArray();
      for (int i = 0; i < m; i++) {
         for (int j = 0; j < n; j++) {
            if(A[i][j]!=0 | B.A[i][j]!=0)
               C[i][j] = A[i][j] - B.A[i][j];
         }
      }
      return X;
   }
   
   /**  Unary minus
   @return    -A
   */
   public Matriz uminus () {
      Matriz X = new Matriz( m, n );
      float[][] C = X.getArray();
      for ( int i = 0; i < m; i++ ) {
         for ( int j = 0; j < n; j++ ) {
            if(A[i][j]!=0)
               C[i][j] = -A[i][j];
         }
      }
      return X;
   }
   
   public Matriz InversionRapida () {
    	
      Matriz J = new Matriz(A);
      Matriz JInv = new Matriz( m, n );
      int m1, m2; //dimensiones de las submatrices cuadradas
      	
      if (m==1){
      JInv.set(0,0,1/A[0][0]);            		
      }
      else{
         m1 = (int) Math.ceil( m/2 );
         m2 = m - m1;
         	
         Matriz P = J.ObtenerSubMatriz(0,m1-1,0,m1-1);
         Matriz Q = J.ObtenerSubMatriz(0,m1-1,m1,m-1);
         Matriz R = J.ObtenerSubMatriz(m1,m-1,0,m1-1);
         Matriz S = J.ObtenerSubMatriz(m1,m-1,m1,m-1);
         
         Matriz K1   = P.InversionRapida();
         Matriz K2   = R.ProductoD(K1);
         Matriz K3   = K1.ProductoD(Q);
         Matriz K4   = R.ProductoD(K3);
         Matriz K5   = K4.minus(S);
         Matriz K6   = K5.InversionRapida();
         Matriz QInv = K3.ProductoD(K6);
         Matriz RInv = K6.ProductoD(K2);
         Matriz K7   = K3.ProductoD(RInv);
         Matriz PInv = K1.minus(K7);
         Matriz SInv = K6.uminus();
         
         JInv.SetearSubMatriz(0,m1-1,0,m1-1,PInv);
         JInv.SetearSubMatriz(0,m1-1,m1,m-1,QInv);
         JInv.SetearSubMatriz(m1,m-1,0,m1-1,RInv);
         JInv.SetearSubMatriz(m1,m-1,m1,m-1,SInv);         
      }      	
      return JInv;  	
   }
   
   static public void MPrint(Matriz AImprimir){
   	int dim1=AImprimir.numeroFil();
   	int dim2=AImprimir.numeroCol();
	String salida="";
	DecimalFormat DosDecimales=new DecimalFormat("0.000");    	
	for(int i=0;i<dim1;i++){
		for(int j=0;j<dim2;j++){
			salida+=DosDecimales.format(AImprimir.get(i,j)) + "  ";				
		}
		salida+="\n";		
	}    	
	System.out.println(salida);   	
   }
   
   public Matriz EliminarFila(int fila){
   		Matriz X=new Matriz(m-1,n);
   		float[][] C=X.getArray();
   		for(int i=0;i<m;i++){
   			for(int j=0;j<n;j++){
   				if(i<fila)
   					C[i][j]=A[i][j];
   				else if(i>fila)
   					C[i-1][j]=A[i][j];
   			}
   		}
   		return X;
   }
   
   public Matriz EliminarColumna(int col){
   		Matriz X=new Matriz(m,n-1);
   		float[][] C=X.getArray();
   		for(int i=0;i<m;i++){
   			for(int j=0;j<n;j++){
   				if(j<col)
   					C[i][j]=A[i][j];
   				else if (j>col)
   					C[i][j-1]=A[i][j];
   			}
   		}
   		return X;
   }
   
   public Matriz InsertarCerosFila(int fila){
   		Matriz X=new Matriz(m+1,n);
   		float[][] C=X.getArray();
   		for(int i=0;i<m+1;i++){
   			for(int j=0;j<n;j++){
   				if(i<fila)
   					C[i][j]=A[i][j];
   				else if(i==fila)
   					C[i][j]=0;
   				else if(i>fila)
   					C[i][j]=A[i-1][j];
   			}
   		}
   		return X;
   }
   
   public Matriz InsertarCerosColumna(int col){
   		Matriz X=new Matriz(m,n+1);
   		float[][] C=X.getArray();
   		for(int i=0;i<m;i++){
   			for(int j=0;j<n+1;j++){
   				if(j<col)
   					C[i][j]=A[i][j];
   				else if(j==col)
   					C[i][j]=0;
   				else if (j>col)
   					C[i][j]=A[i][j-1];
   			}
   		}
   		return X;
   }      
   
   /* ------------------------
   Private Methods
 * ------------------------ */

   /** Check if size(A) == size(B) **/

   private void checkMatrixDimensions (Matriz B) {
      if (B.m != m || B.n != n) {
         throw new IllegalArgumentException("Matrix dimensions must agree.");
      }
   }   
      
}