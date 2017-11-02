import java.util.ArrayList;

public class demo {
	public static void main(String[] args){
		ArrayList list = new ArrayList();
		list.add("11");
		list.add("22");
		list.add("33");
		for(int i=0;i<list.size();i++){
			System.out.println(list.get(i));
		}
	}

}
