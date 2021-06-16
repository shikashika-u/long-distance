//解析したい動画があるフォルダを選択
opendir = getDirectory("入力フォルダを選択")
//結果を出力するフォルダを選択
savedir = getDirectory("出力フォルダを選択")
moviedir = savedir+"\\movie"
ROIimagedir = savedir+"\\ROIimage"
Resultdir = savedir+"\\Result"
ROI = getNumber("ROI数を入力", 2)
//保存先のフォルダ作成
File.makeDirectory(moviedir);
File.makeDirectory(ROIimagedir);
File.makeDirectory(Resultdir);
//ファルダ内のファイル名リストを参照元ディレクトリから配列で取得
filelist = getFileList(opendir);
//Result table上の数値をクリア
run("Clear Results");
//フォルダにある動画ファイルを二つずつ開いていく(RFP⇒GFPの順に撮影されている前提)
for (k=0; k<filelist.length; k=k+2){
	open(opendir+"\\"+filelist[k]);
	rename("2");
	open(opendir+"\\"+filelist[k+1]);
	rename("1");
	ratioimage();//ImageCalculatorとzprojectで画像処理
	for (i=0; i<=ROI-1; i++){
		waitForUser("ROI selection", "ROIを選択");
		roiManager("add");
	}//ROIを手動で選択するとROImanagerに登録してくれる。
	selectWindow("Result of Result of Result of 1");
	roiManager("Select", 0);
	run("Plot Z-axis Profile");
	Plot.getValues(x, y);
	for (i=0; i<x.length; i++){
		setResult(filelist[k+1]+":time", i, x[i]);
		setResult(filelist[k+1]+":ROI1", i, y[i]);
	}
	close();
	for (i=1; i<=ROI-1; i++) {
		ROInumber = filelist[k+1]+":ROI"+i+1;
		roiManager("select", i);
		run("Plot Z-axis Profile");
		Plot.getValues(x, y);
		for(j=0; j<x.length; j++){
			setResult(ROInumber, j, y[j]); 
		}
		close();
	}
	selectWindow("Result of Result of Result of 1");
	close();
	ROIimagesave(k);
	moviesave(k);
	run("Close All");
}

//ratio画像を作成ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
function ratioimage(){
	selectWindow("1");
	run("Z Project...", "stop=20 projection=[Average Intensity]");
	imageCalculator("Subtract create 32-bit stack", "1","AVG_1");
	imageCalculator("Divide create 32-bit stack", "Result of 1","AVG_1");
	selectWindow("Result of 1");
	close();
	selectWindow("2");
	run("Z Project...", "projection=[Average Intensity]");
	imageCalculator("Divide create 32-bit stack", "Result of Result of 1","AVG_2");
	selectWindow("AVG_1");
	close();
	selectWindow("AVG_2");
	close();
	selectWindow("Result of Result of 1");
	close();
}

//ROI情報入りの画像を保存ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
function ROIimagesave(a){
	selectWindow("1");
	run("Z Project...", "projection=[Average Intensity]");
	setMinAndMax(0, 20);
	run("Apply LUT");
	run("Scale Bar...", "width=3000 height=10 font=30 color=White background=None location=[Lower Right] bold overlay");
	roiManager("deselect");
	roiManager("Set Line Width", 5);
	roiManager("Show All");
    Overlay.flatten
	roiManager("deselect");
	roiManager("delete");
	saveAs("Jpeg", ROIimagedir+"\\"+filelist[a]);
	close();
	selectWindow("AVG_1");
	close();
}

//動画を軽量化して保存ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
function moviesave(a){
	selectWindow(1);
	makeLine(0, 0, 0, 0);
	run("Size...", "width=400 height=600 constrain average interpolation=Bilinear");
	saveAs("Tiff", moviedir+"\\"+filelist[a+1]);
	selectWindow("2");
	run("Size...", "width=400 height=400 constrain average interpolation=Bilinear");
	saveAs("Tiff", moviedir+"\\"+filelist[a]);
}
	
