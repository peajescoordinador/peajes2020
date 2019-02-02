/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package cl.coordinador.peajes;

import javax.swing.DefaultCellEditor;
import javax.swing.JComboBox;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;

/**
 * Interfaz de usuario para visualizacion y modificacion de las opciones de
 * configuracion de la herramienta
 *
 * @author Frank Leanez at www.flconsulting.cl
 */
public class PeajesConfigGUI extends javax.swing.JDialog {
    
    private java.util.Properties propiedades = null;
    private KeyPropertyTable tblMain;
    
    /**
     * Crea una nueva interfaz de usuario para visualizacion y modificacion de
     * las opciones de configuracion de la herramienta
     * <br>Recordar llamar setVisible(true) para desplegar al usuario
     *
     * @param parent Ventana padre (acepta usar clase PeajesCDEC porque extiende
     * la clase JFrame
     * @param propiedades instacia de propiedades. Debe estar incializada! No
     * usar null!
     */
    public PeajesConfigGUI(javax.swing.JFrame parent, java.util.Properties propiedades) {
        super(parent, true);
        initComponents();
        initTable();
        assert (propiedades != null) : "No dijimos que no se permitia null aqui??";
        this.propiedades = propiedades;
        loadProperties(this.propiedades);
    }
    
    private void initTable () {
        tblMain = new KeyPropertyTable();

        tblMain.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Campo", "Valor", "Editable"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, KeyPropertyValue.class
            };
            boolean[] canEdit = new boolean [] {
                false, true, true
            };

            @Override
            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            @Override
            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        tblMain.getTableHeader().setReorderingAllowed(false);
        tblMain.getColumnModel().getColumn(2).setMinWidth(0);
        tblMain.getColumnModel().getColumn(2).setMaxWidth(0);
        tblMain.getColumnModel().getColumn(2).setWidth(0);
        javax.swing.JScrollPane jScrollPane12 = new javax.swing.JScrollPane();
        jScrollPane12.setViewportView(tblMain);
        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane12, javax.swing.GroupLayout.DEFAULT_SIZE, 400, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane12, javax.swing.GroupLayout.DEFAULT_SIZE, 286, Short.MAX_VALUE)
        );
    }
    
    private void loadProperties (java.util.Properties propiedades) {
        java.util.Map<Object, Object> sortedPropiedades = new java.util.TreeMap(propiedades);
        for (java.util.Map.Entry<Object, Object> o: sortedPropiedades.entrySet()){
            String keyRaw = o.getKey().toString().trim();
            String valRaw = o.getValue().toString().trim();
            KeyPropertyValue key = new KeyPropertyValue(keyRaw);
            tblMain.addRow(key, valRaw);
        }
    }
    
    private void userWantsExit() {
        int yesNoCancel = JOptionPane.showConfirmDialog(this, "¿Desea guardar los cambios?", "Salir", JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE);
        switch (yesNoCancel) {
            case JOptionPane.YES_OPTION:
                userWantsSave();
                dispose();
                break;
            case JOptionPane.NO_OPTION:
                dispose();
                break;
            case JOptionPane.CANCEL_OPTION:
                break;
        }
    }
    
    private void userWantsSave() {
        tblMain.updateProperties(propiedades);
        try {
            PeajesCDEC.saveOptionFile(propiedades);
            
            System.out.println("Guardado de propiedades de configuracion exitosa");
        } catch (java.io.FileNotFoundException e) {
            JOptionPane.showMessageDialog(this, "Error interno. No se pudo grabar el archivo", "Error guardad", JOptionPane.WARNING_MESSAGE);
            e.printStackTrace(System.err);
            userWantsExit();
        } catch (java.io.IOException ex) {
            int yesNoCancel = JOptionPane.showConfirmDialog(this, "Error al guardar el archivo " + ex.getMessage() + "\n¿Desea intentar nuevamente?", "Salir", JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE);
            switch (yesNoCancel) {
                case JOptionPane.YES_OPTION:
                    userWantsSave();
                    break;
                case JOptionPane.NO_OPTION:
                    userWantsExit();
                    break;
                case JOptionPane.CANCEL_OPTION:
                    break;
            }
        }
    }


    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        menuMainBar = new javax.swing.JMenuBar();
        menuFile = new javax.swing.JMenu();
        menuSave = new javax.swing.JMenuItem();
        menuExit = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.DO_NOTHING_ON_CLOSE);
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowClosing(java.awt.event.WindowEvent evt) {
                formWindowClosing(evt);
            }
        });

        menuFile.setText("Archivo");

        menuSave.setText("Guardar");
        menuSave.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                menuSaveActionPerformed(evt);
            }
        });
        menuFile.add(menuSave);

        menuExit.setText("Salir");
        menuExit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                menuExitActionPerformed(evt);
            }
        });
        menuFile.add(menuExit);

        menuMainBar.add(menuFile);

        setJMenuBar(menuMainBar);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 400, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 286, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void menuExitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_menuExitActionPerformed
        userWantsExit();
    }//GEN-LAST:event_menuExitActionPerformed

    private void menuSaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_menuSaveActionPerformed
        userWantsSave();
    }//GEN-LAST:event_menuSaveActionPerformed

    private void formWindowClosing(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowClosing
        userWantsExit();
    }//GEN-LAST:event_formWindowClosing


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JMenuItem menuExit;
    private javax.swing.JMenu menuFile;
    private javax.swing.JMenuBar menuMainBar;
    private javax.swing.JMenuItem menuSave;
    // End of variables declaration//GEN-END:variables
}

/**
 * Clase especial de JTable para manejar las opciones que estan en un forma especifico dentro de un mapa tipo java.util.Properties
 * @author Frank Leanez at www.flconsulting.cl
 */
class KeyPropertyTable extends javax.swing.JTable {
    
    /**
     * Sets the value for the cell in the table model at row and column.
     *
     * @param aValue the new value
     * @param row the row of the cell to be changed
     * @param column the column of the cell to be changed
     */
    @Override
    public void setValueAt(Object aValue, int row, int column) {
        //Solo se aceptan cambios en la columna 2 (columna de valores):
        if (column == 1) {
            Object o = getValueAt(row, 2);
            if (o != null) {
                //No aceptamos vacios ni cambios en la llave 'Version'
                if (o.toString().equals("Version") || aValue.toString().isEmpty()) {
                    return;
                }
                
                //Ahora vemos si el valor que quiere ingresar el usuario en la tabla es aceptable (para la validacion):
                KeyPropertyValue key = (KeyPropertyValue) o;
                try {
                    switch (key.getType()) {
                        case BOOLEAN:
                            boolean b = Boolean.parseBoolean(aValue.toString());
                            break;
                        case DOUBLE:
                            double d = Double.parseDouble(aValue.toString());
                            break;
                        case INTEGER:
                            int i = Integer.parseInt(aValue.toString());
                            break;
                    }
                    super.setValueAt(aValue, row, column);
                } catch (NumberFormatException e) {
                    System.out.println("Invalid '" + key.getType().name() + "' " + e.getMessage());
                }
            }
        }
    }

    /**
     * Returns an appropriate editor for the cell specified by row and column.
     * If the TableColumn for this column has a non-null editor, returns that.
     * If not, finds the class of the data in this column (using getColumnClass)
     * and returns the default editor for this type of data.
     * <br><b>Note:</b> Throughout the table package, the internal
     * implementations always use this method to provide editors so that this
     * default behavior can be safely overridden by a subclass.
     *
     * @param row the row of the cell to edit, where 0 is the first row
     * @param column the column of the cell to edit, where 0 is the first column
     * @return the editor for this cell; if null return the default editor for
     * this type of cell
     */
    @Override
    public TableCellEditor getCellEditor(int row, int column) {
        Object o = getValueAt(row, 2);
        if (o instanceof KeyPropertyValue) {
            KeyPropertyValue key = (KeyPropertyValue) o;
            switch (key.getType()) {
                case BOOLEAN:
                    Boolean[] items = new Boolean[]{Boolean.FALSE, Boolean.TRUE};
                    JComboBox combo = new JComboBox(items);
                    return new DefaultCellEditor(combo);
            }
            if (key.hasValidation()) {
                JComboBox combo = new JComboBox(key.getValidation());
                return new DefaultCellEditor(combo);
            }
        }
        return super.getCellEditor(row, column);
    }

    /**
     * Agrega una nueva entrada (fila) a la tabla
     *
     * @param key instancia de la llave especial KeyPropertyValue (no usar
     * null!)
     * @param value valor correspondiente a la llave
     */
    public void addRow(KeyPropertyValue key, String value) {
        DefaultTableModel model = (DefaultTableModel) getModel();
        java.util.Vector row = new java.util.Vector();
        row.addElement(key.getDisplayName());
        row.addElement(value);
        row.addElement(key);
//        String[] row = new String[3];
//        row[0] = key.getDisplayName();
//        row[1] = value;
//        row[2] = key;
        model.addRow(row);
    }

    /**
     * Actualiza la instancia de propiedades con los valores mostrados en la
     * tabla
     *
     * @param propiedades instancia de propiedades. no use null!
     */
    public void updateProperties(java.util.Properties propiedades) {
        for (int row = 0; row < getRowCount(); row++) {
            String key = getValueAt(row, 2).toString();
            String value = getValueAt(row, 1).toString();
            propiedades.put(key, value);
        }
    }
}

/**
 * Clase auxiliar para el manejo de esos "key" especiales del archivo de opciones
 * @author Frank Leanez at www.flconsulting.cl
 */
class KeyPropertyValue {
    private final String displayName;
    private final String keyRaw;
    private PeajesConstant.DataType type;
    private java.util.List<String> validation;
    
    public KeyPropertyValue(String keyRaw) {
        this.keyRaw = keyRaw;
        String[] keys = keyRaw.split(PeajesConstant.KEY_SEPARATOR);
        if (keys.length == 1) {
            displayName = keys[0];
            type = PeajesConstant.DataType.STRING;
            validation = new java.util.ArrayList<String>();
        } else if (keys.length == 2) {
            displayName = keys[0];
            type = PeajesConstant.DataType.valueOf(keys[1]);
            validation = new java.util.ArrayList<String>();
        } else if (keys.length > 2) {
            displayName = keys[0];
            type = PeajesConstant.DataType.valueOf(keys[1]);
            validation = new java.util.ArrayList<String>();
            for (int i = 2; i < keys.length; i++) {
                validation.add(keys[i]);
            }
        } else {
            displayName = keyRaw;
            type = PeajesConstant.DataType.STRING;
            validation = new java.util.ArrayList<String>();
        }
    }
    
    public Class getClassKey () {
        switch (type) {
            case BOOLEAN:
                return Boolean.class;
            case DOUBLE:
                return Double.class;
            case INTEGER:
                return Integer.class;
            case STRING:
                return String.class;
            default:
                return String.class;
        }
    }

    public String getDisplayName() {
        return displayName;
    }

    public PeajesConstant.DataType getType() {
        return type;
    }

    public String[] getValidation() {
        return validation.toArray(new String[validation.size()]);
    }
    
    public boolean hasValidation () {
        return !validation.isEmpty();
    }
    
    @Override
    public String toString() {
        return keyRaw;
    }
    
}