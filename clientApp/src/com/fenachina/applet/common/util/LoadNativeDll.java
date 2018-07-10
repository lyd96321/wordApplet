package com.fenachina.applet.common.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class LoadNativeDll {
	
	/**
	 * 初始化动态链接库
	 * @throws IOException
	 */
    public static void loadLib() throws IOException {  
        String systemType = System.getProperty("os.name"); 
        String arch=System.getProperty("sun.arch.data.model");
        String libName="/";
        String jacobDllName="";
        String jacobRealDllName="";
        String libExtension = (systemType.toLowerCase().indexOf("win")!=-1) ? ".dll" : ".so";  
        if("32".equals(arch)){
        	jacobDllName+="jacob1-18x86";
        	jacobRealDllName="jacob-1.18-x86";
        }else{
        	jacobDllName+="jacob1-18x64";
        	jacobRealDllName="jacob-1.18-x64";
        } 
        jacobDllName=jacobDllName+libExtension;
        jacobRealDllName+=libExtension;
        String libFullName = libName+jacobDllName; 
        String nativeTempDir = System.getProperty("java.io.tmpdir");  
        File extractedLibFile = new File(nativeTempDir+File.separator+jacobRealDllName);
        String libPath=System.getProperty("java.home")+File.separator+"bin";
        File libFile = new File(libPath+File.separator+jacobRealDllName);
        //写jacob动态文件
        if(!libFile.exists()){
        	writeDllToJre(libFullName,libFile);
        }
        if(!extractedLibFile.exists()){  
        	//将dll文件写入  临时缓存中
        	writeDllToJre(libName+jacobDllName,extractedLibFile);
	        try{
	        	System.load(extractedLibFile.toString()); 
	        }catch(Exception e){
	        	
	        }
        }
    }
    
    /**
     * 将动态文件写入可读位置：jre环境变量中、temp高速缓存中
     * @param libFullName
     * @param libFile
     */
    private static void writeDllToJre(String libFullName,File libFile){
    	 InputStream in = null;  
         BufferedInputStream reader = null;  
         FileOutputStream writer = null;  
    	 if(!libFile.exists()){  
             try {  
                 in = LoadNativeDll.class.getClassLoader().getResourceAsStream(libFullName);  
                 if(in==null){
                     in =  LoadNativeDll.class.getResourceAsStream(libFullName); 
                 } 
                 reader = new BufferedInputStream(in);  
                 writer = new FileOutputStream(libFile);  
                 byte[] buffer = new byte[1024];  
                 while (reader.read(buffer) > 0){  
                     writer.write(buffer);  
                     buffer = new byte[1024];  
                 }  
                 writer.flush();
                 if(in!=null)  
                      in.close();
				if(writer!=null)
				     writer.close();
				if(reader!=null)
					 reader.close();
             } catch (IOException e){  
                 e.printStackTrace();  
             } finally {  
					try {
						if(in!=null)  
		                     in.close();
						if(writer!=null)
						writer.close();
						if(reader!=null)
							reader.close();
					} catch (IOException e) {
						e.printStackTrace();
					}  
             }  
         }  
    }

}
