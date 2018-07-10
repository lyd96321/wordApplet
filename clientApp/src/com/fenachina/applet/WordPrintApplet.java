package com.fenachina.applet;

import java.applet.Applet;
import java.io.IOException;
import java.security.AccessController;
import java.security.PrivilegedAction;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import javax.swing.SwingUtilities;

import com.fenachina.applet.common.util.LoadNativeDll;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
/**
 * word打印程序
 * @Title: Controller  
 * @Description: TODO
 * @date 2018年6月14日
 * @version V1.0   
 *
 */
public class WordPrintApplet extends Applet {
    /**
     * 
     */
    private static final long serialVersionUID = 6293078385198700265L;


    public void init() {
        // 创建applet的GUI，由event dispatching thread运行
        try {
        	 SwingUtilities.invokeAndWait(new Runnable() {
                 public void run() {
                	 try {
         				LoadNativeDll.loadLib();
         			} catch (IOException e) {
         			}
                	 
                 }
             });
        } catch (Exception e) {
        	e.printStackTrace();
            System.err.println("createGUI didn't complete successfully");
        }
    }

    //得到要打印的pdf文件的url列表，并开始打印
    public void launchPrint(final String url) {
    	 AccessController.doPrivileged(new PrivilegedAction<Object>() {
            public Object run() {
            	if(url==null){
    		      return null;
    	         }
            	 List<String> inputArr = convertArrayParameter(url);
                 try {
          		    print(inputArr);
          	     } catch (IOException e) {
          		  e.printStackTrace();
          	  }
                 return null;
         }
            
    	});
    }


    // 解析传给Applet的参数，将以逗号分隔的字符串分成参数List
    private List<String> convertArrayParameter(String parameterArrStr) {
        List<String> list = new ArrayList<String>();
        String[] parameterArr = parameterArrStr.split(",");
        Collections.addAll(list, parameterArr);
        return list;
    }



    //执行打印
    private void print(List<String> pdfURLList) throws IOException {
        for (String wordURL : pdfURLList) {
        	printWord(wordURL);
        }
    }
    
    
    private void printWord(String url){
    	    ComThread.InitSTA();
	        ActiveXComponent word=new ActiveXComponent("Word.Application");
	        Dispatch doc=null;
	        Dispatch.put(word, "Visible", new Variant(false));
	        Dispatch docs=word.getProperty("Documents").toDispatch();
	        if(url!=null&&!url.trim().equals("")){
	        	 doc=Dispatch.call(docs, "Open",url).toDispatch();
	        	 Dispatch.call(doc, "PrintOut");//打印
	        }
	        try {
	                if(doc!=null){
	                    Dispatch.call(doc, "Close",new Variant(0));
	                }
	                //关闭进程
		            if(word!=null){
		            	word.invoke("Quit",new Variant[]{});
		            }
		          //释放资源
		            ComThread.Release();
	            } catch (Exception e2) {
	                e2.printStackTrace();
	            }finally {
	            	  if(doc!=null){
		                    Dispatch.call(doc, "Close",new Variant(0));
		                }
		                //关闭进程
			            if(word!=null){
			            	word.invoke("Quit",new Variant[]{});
			            }
			          //释放资源
			            ComThread.Release();
				}
	         
    }
    
    
    
  
}